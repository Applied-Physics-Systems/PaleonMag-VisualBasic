VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmADWIN_AF 
   Caption         =   "ADWIN AF Ramp"
   ClientHeight    =   8295
   ClientLeft      =   255
   ClientTop       =   4200
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   6855
   Begin VB.Frame Frame3 
      Caption         =   "AF Coil Temperature"
      Height          =   1575
      Left            =   2520
      TabIndex        =   43
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdTemp 
         Caption         =   "Refresh T"
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtTemp2 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtTemp1 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblAFtooHot 
         Caption         =   "The AF unit is too hot so let's pause a little bit..."
         Height          =   615
         Left            =   2520
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "°C"
         Height          =   255
         Left            =   2040
         TabIndex        =   47
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "°C"
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Transver Coil:"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Axial Coil:"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shiny Buttons"
      Height          =   3135
      Left            =   120
      TabIndex        =   37
      Top             =   4440
      Width           =   2295
      Begin VB.CommandButton cmdGotoAFSettings 
         Caption         =   "Open AF Settings"
         Height          =   372
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdStartAFTuner 
         BackColor       =   &H00FF80FF&
         Caption         =   "Tune AF Coils"
         Height          =   372
         Left            =   240
         MaskColor       =   &H00C000C0&
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdTestGaussMeter 
         Caption         =   "908A Gaussmeter Control"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CommandButton cmdCalibrate 
         Caption         =   "Calibrate AF Coils"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdChangeFileSaveSettings 
         Caption         =   "Open AF File Settings"
         Height          =   372
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "AF Ramp Mode"
      Height          =   2415
      Left            =   120
      TabIndex        =   36
      Top             =   1920
      Width           =   2295
      Begin VB.CheckBox chkDCFieldRecord 
         Caption         =   "Record DC Field"
         Height          =   372
         Left            =   240
         TabIndex        =   52
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdCleanCoils 
         Caption         =   "Clean Coils"
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox chkVerbose 
         Caption         =   "Debug Mode?"
         Height          =   372
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkClippingTest 
         Caption         =   "Unmonitored Ramp"
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Active Coil"
      Height          =   1695
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton SWITCH_BUTTON 
         Caption         =   "Switch"
         Height          =   315
         Left            =   1440
         TabIndex        =   54
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "IRM Axial"
         Height          =   192
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkLockCoils 
         Caption         =   "Lock coil selection"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Transverse"
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1212
      End
      Begin VB.OptionButton optCoil 
         Caption         =   "Axial"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "AF Ramp Setup"
      Height          =   6375
      Left            =   2520
      TabIndex        =   25
      Top             =   1800
      Width           =   4215
      Begin VB.CheckBox chkAutoRampSlope 
         Caption         =   "Calculate Ramp Slopes Automatically"
         Height          =   435
         Left            =   360
         TabIndex        =   50
         Top             =   3600
         Width           =   2175
      End
      Begin VB.OptionButton optCalRamp 
         Caption         =   "Uncalibrated Ramp (use Peak Monitor Voltage)"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   2640
         Width           =   3735
      End
      Begin VB.OptionButton optCalRamp 
         Caption         =   "Calibrated Ramp (use Peak Field value)"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   3000
         Width           =   3375
      End
      Begin VB.ComboBox cmbFieldUnits 
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Text            =   "G"
         Top             =   1560
         Width           =   852
      End
      Begin VB.TextBox txtPeakField 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   1560
         Width           =   972
      End
      Begin VB.TextBox txtFreq 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   972
      End
      Begin MSComDlg.CommonDialog cdlgTestSineFit 
         Left            =   3600
         Top             =   480
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
      End
      Begin VB.TextBox txtMonitorTrigVolt 
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   1920
         Width           =   972
      End
      Begin VB.TextBox txtRampRate 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   1080
         Width           =   972
      End
      Begin VB.TextBox txtRampDownSlope 
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   4800
         Width           =   972
      End
      Begin VB.TextBox txtRampUpSlope 
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Top             =   4320
         Width           =   972
      End
      Begin VB.CommandButton cmdStartRamp 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Start Ramp"
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   5760
         Width           =   3255
      End
      Begin VB.TextBox txtRampPeakDuration 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   720
         Width           =   972
      End
      Begin VB.TextBox txtRampPeakVoltage 
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   2280
         Width           =   972
      End
      Begin VB.Label lblTotalRampDuration 
         Caption         =   "Label5"
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Total Ramp Time (ms):"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label lblRampDownDuration 
         Caption         =   "Label4"
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label lblRampUpDuration 
         Caption         =   "Label4"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Duration (ms):"
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Units:"
         Height          =   255
         Left            =   3120
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Peak Field:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Sine Freq. (Hz):"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label33 
         Caption         =   "Peak Monitor Voltage:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label23 
         Caption         =   "Ramp IO Rate (Hz):"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label22 
         Caption         =   "Ramp Down Slope (volts/sec):"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Ramp Up Slope (volts/sec):"
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Time at Peak (ms):"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Peak Ramp Voltage:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H008080FF&
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000C0&
      TabIndex        =   9
      Top             =   7680
      Width           =   2295
   End
End
Attribute VB_Name = "frmADWIN_AF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RandomCoilClick As Boolean
Dim isUserChange As Boolean

Private ramp_in_progress As Boolean

Private Sub chkAutoRampSlope_Click()

    If Me.chkAutoRampSlope.value = Checked Then
    
        'Disable the Ramp Up & Ramp Down controls
        Me.txtRampUpSlope.Enabled = False
        Me.txtRampDownSlope.Enabled = False

    Else
    
        'Enable the Ramp Up & Ramp Down controls
        Me.txtRampUpSlope.Enabled = True
        Me.txtRampDownSlope.Enabled = True

    End If

End Sub

Private Sub chkClippingTest_Click()

    'if user checks this, then need to disable the Peak Field
    'Peak Monitor Voltage and optCalRamp controls
    If chkClippingTest.value = Checked Then
    
        Me.txtPeakField.Enabled = False
        Me.txtMonitorTrigVolt.Enabled = False
        Me.optCalRamp(0).Enabled = False
        Me.optCalRamp(1).Enabled = False

    Else
    
        'User wants a calibrated ramp, enable all the controls
        Me.txtPeakField.Enabled = True
        Me.txtMonitorTrigVolt.Enabled = True
        Me.optCalRamp(0).Enabled = True
        Me.optCalRamp(1).Enabled = True
        
    End If

End Sub

Private Sub chkLockCoils_Click()

    If Me.chkLockCoils.value = Checked Then
    
        CoilsLocked = True
        optCoil(0).Enabled = False
        optCoil(1).Enabled = False
        
    ElseIf Me.chkLockCoils.value = Unchecked Then
    
        CoilsLocked = False
        optCoil(0).Enabled = True
        optCoil(1).Enabled = True
        
    End If

End Sub

Public Sub CleanCoils()

    modAF_DAQ.ClearAFErrorStatus
    modAF_DAQ.InitDefault_AFErrorResponseTable
    
    'Suppress CodeYellow and CodeRed Target Undershoot errors
    modAF_DAQ.AF_Error_Response_Table(TargetUndershoot, coil_type.Axial, CodeYellow).ErrorAction = SuppressError
    modAF_DAQ.AF_Error_Response_Table(TargetUndershoot, coil_type.Transverse, CodeYellow).ErrorAction = SuppressError
    modAF_DAQ.AF_Error_Response_Table(TargetUndershoot, coil_type.Axial, CodeRed).ErrorAction = SuppressError
    modAF_DAQ.AF_Error_Response_Table(TargetUndershoot, coil_type.Transverse, CodeRed).ErrorAction = SuppressError

    ExecuteRamp AxialCoilSystem, _
                AfAxialMax, _
                , , , _
                0, _
                True, _
                False, _
                (Me.chkVerbose.value = Checked)
                
    ExecuteRamp TransverseCoilSystem, _
                AfTransMax, _
                , , , _
                0, _
                True, _
                False, _
                (Me.chkVerbose.value = Checked)
                
    'Set AF Error response table back to default
    modAF_DAQ.ClearAFErrorStatus
    modAF_DAQ.InitDefault_AFErrorResponseTable
                
End Sub

Private Sub cmdCalibrate_Click()

    frmCalibrateCoils.InAFMode = True
    Load frmCalibrateCoils
    frmCalibrateCoils.Show
        
End Sub

Private Sub cmdChangeFileSaveSettings_Click()

    Load frmFileSave
    frmFileSave.ZOrder
    frmFileSave.Show

End Sub

Private Sub cmdCleanCoils_Click()

    CleanCoils

End Sub

Private Sub cmdClose_Click()
    
    Me.Hide
    
End Sub

Private Sub cmdGotoAFSettings_Click()

    'Load the settings form
    Load frmSettings
    
    'Select the AF settings tab
    frmSettings.selectTab 4
    
    'Show the Settings form
    frmSettings.Show
    
End Sub

Private Sub cmdOpenDAQComm_Click()

    frmDAQ_Comm.Show
    
End Sub

Private Sub cmdStartAFTuner_Click()
    
    Me.Hide
    frmAFTuner.Show
            
End Sub

Public Sub cmdStartRamp_Click()

    modAF_DAQ.ClearAFErrorStatus
    modAF_DAQ.InitDefault_AFErrorResponseTable

    If Me.chkClippingTest.value = Checked Then
    
        If Me.chkAutoRampSlope.value = Checked Then
        
            'Don't put in up and down slopes
            ExecuteRamp ActiveCoilSystem, _
                        val(Me.txtRampPeakVoltage), , , _
                        val(Me.txtRampRate), _
                        val(Me.txtRampPeakDuration), _
                        False, _
                        True, _
                        (Me.chkVerbose.value = Checked), _
                        (Me.chkDCFieldRecord.value = Checked)
                        
        Else
        
            'Need to put in the Ramp Up & Down slopes
            ExecuteRamp ActiveCoilSystem, _
                        val(Me.txtRampPeakVoltage), _
                        val(Me.txtRampUpSlope), _
                        val(Me.txtRampDownSlope), _
                        val(Me.txtRampRate), _
                        val(Me.txtRampPeakDuration), _
                        False, _
                        True, _
                        (Me.chkVerbose.value = Checked), _
                        (Me.chkDCFieldRecord.value = Checked)
        
        End If
                    
    ElseIf Me.optCalRamp(1).value = True Then
    
        If Me.chkAutoRampSlope.value = Checked Then
        
            'Don't put in up and down slopes
            ExecuteRamp ActiveCoilSystem, _
                        val(Me.txtMonitorTrigVolt), , , _
                        val(Me.txtRampRate), _
                        val(Me.txtRampPeakDuration), _
                        False, _
                        False, _
                        (Me.chkVerbose.value = Checked), _
                        (Me.chkDCFieldRecord.value = Checked)
                        
        Else
        
            'Need to put in the Ramp Up & Down slopes
            ExecuteRamp ActiveCoilSystem, _
                        val(Me.txtMonitorTrigVolt), _
                        val(Me.txtRampUpSlope), _
                        val(Me.txtRampDownSlope), _
                        val(Me.txtRampRate), _
                        val(Me.txtRampPeakDuration), _
                        False, _
                        False, _
                        (Me.chkVerbose.value = Checked), _
                        (Me.chkDCFieldRecord.value = Checked)
        
        End If
                    
    ElseIf Me.optCalRamp(0).value = True Then
        
        If Me.chkAutoRampSlope.value = Checked Then
        
            'Don't put in up and down slopes
            ExecuteRamp ActiveCoilSystem, _
                        val(Me.txtPeakField), , , _
                        val(Me.txtRampRate), _
                        val(Me.txtRampPeakDuration), _
                        True, _
                        False, _
                        (Me.chkVerbose.value = Checked), _
                        (Me.chkDCFieldRecord.value = Checked)
                        
        Else
        
            'Need to put in the Ramp Up & Down slopes
            ExecuteRamp ActiveCoilSystem, _
                        val(Me.txtPeakField), _
                        val(Me.txtRampUpSlope), _
                        val(Me.txtRampDownSlope), _
                        val(Me.txtRampRate), _
                        val(Me.txtRampPeakDuration), _
                        True, _
                        False, _
                        (Me.chkVerbose.value = Checked), _
                        (Me.chkDCFieldRecord.value = Checked)
        
        End If

    Else
    
        'Wha?
        MsgBox "Whoops!"
                    
    End If
          
End Sub

Private Sub cmdTemp_Click()

    '(July 2010 - I Hilburn)
    'Copied from Laurent's code in frmAF_2G. Only Changed textbox object references.

    Dim Temp1 As Double ' (February 2010 L Carporzen) Monitor temperature of the coils before executing AF
    Dim Temp2 As Double
    
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    If EnableT1 Then
    
        Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
        
        Temp1 = Temp1 * TSlope - Toffset
        
    End If
        
    txtTemp1 = Format$(Temp1, "##0.00")
    
    If EnableT2 Then
    
        Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
        
        Temp2 = Temp2 * TSlope - Toffset
        
    End If
        
    txtTemp2 = Format$(Temp2, "##0.00")
    
End Sub

Private Sub cmdTestGaussMeter_Click()

    frm908AGaussmeter.Show

End Sub

Public Function DoRampADWIN(ByRef MonitorWave As Wave, _
                            ByRef UpWave As Wave, _
                            ByRef DownWave As Wave, _
                            ByRef AF_Data() As Double, _
                            ByVal PeakField As Double, _
                            Optional ByVal HangeTime As Long = 0, _
                            Optional ByVal RampMode As Long = 1, _
                            Optional ByVal RampDownMode As Long = 0, _
                            Optional ByVal DoDCFieldRecord As Boolean = False) As Long
                       
                       
    DoRampADWIN = DoRampADWIN_WithParameterLogging(MonitorWave, _
                                                   UpWave, _
                                                   DownWave, _
                                                   AF_Data, _
                                                   PeakField, _
                                                   HangeTime, _
                                                   RampMode, _
                                                   RampDownMode, _
                                                   DoDCFieldRecord)
                       
    Exit Function
                       
    Dim ReturnVal As Long
    Dim ProcessFile As String
    Dim ProcessName As String
    Dim PercentDone As String
    
    Dim i As Long
    Dim N As Long
    Dim TempDataIn() As Long
    Dim TempDataOut() As Long
    
    Dim TempD As Double
    Dim TempS As String
    Dim TempL As Long
    Dim Temp1 As Double
    Dim Temp2 As Double
    Dim TWarning As Boolean
    
    Dim StartTime As Long
    Dim RampDuration As Long
    Dim UserResp As Long
    Dim NumPeriods As Long
        
    Dim ErrorMessage As String
    
    Dim DCFieldWave As Wave
    Dim DCFieldStatus As Boolean
    
    'Exit if nocomm mode
    If NOCOMM_MODE = True Then
    
        'This is not an error free ramp
        DoRampADWIN = -616
        
        Exit Function
        
    End If
    
    '(July 23, 2010 - I Hilburn)
    'Added in check to see if the program flow is paused.  If so, the AF code will wait until the flow
    'is returned to Resume
    If Prog_paused = True Then
    
        'Set TempS equal to the current contents
        'of the 2nd status bar panel
        TempS = frmProgram.sbStatusBar.Panels(2).text
        
        'Update the program form status bar
        frmProgram.StatusBar "Paused...", 2
        
        'Loop and wait for status to change and DoEvents
        'to allow user to make changes
        Do
        
            'Pause 100 ms
            PauseTill timeGetTime() + 100
            
            DoEvents
            
        Loop Until Prog_paused = False
        
        'Return the old string to panel 2 of the status bar
        frmProgram.StatusBar TempS, 2
        
    End If
        
    'Do a double check real quick of the max ramp & monitor voltages
    If (ActiveCoilSystem = AxialCoilSystem And _
        (modConfig.AfAxialRampMax = -1 Or _
         modConfig.AfAxialMonMax = -1) And _
        RampMode <> 3) Or _
       (ActiveCoilSystem = TransverseCoilSystem And _
        (modConfig.AfTransRampMax = -1 Or _
         modConfig.AfTransMonMax = -1) And _
        RampMode <> 3) _
    Then
    
        'Send an email and pop-up an error message
        SetCodeLevel CodeRed
        
        'Pause the flow
        Flow_Pause
        
        'Send the error notification
        frmSendMail.MailNotification "AF Settings Error", _
                                     "The Maximum Ramp voltages allowed for the ADWIN AF system " & _
                                     "have not been set yet.  Please perform an AF Clipping-Test now.", _
                                     CodeRed, _
                                     True
                                     
        MsgBox "The Maximum Ramp voltages allowed for the ADWIN AF system " & _
               "have not been set yet.  Please perform an AF Clipping-Test now.", _
               vbOK, _
               "AF Settings Error"
    
        Load frmAFTuner
        frmAFTuner.Show
        
        Me.Hide
        
        DoRampADWIN = -616
        
        Exit Function
        
    End If
        
    'Update status bar to show that the AF run is being configured
    frmProgram.StatusBar "AF Config", 3
        
'------------------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------------------'
'
'   July 2010
'   Authors:
'   Laurent Corporozen
'   Isaac Hilburn
'
'   Copied temperature check code from frmAF_2G.ExecuteRamp to this ADWIN central ramp function
'   Code copied verbatim with minor changes to switch comm implementation from frmMCC to frmDAQ_Comm
'   frmDAQ_Comm takes Channel objects instead of channel port numbers
'------------------------------------------------------------------------------------------------------------------------------------'
        
    'Before doing anything with the ADWIN board, get the AF coil temperatures
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    If EnableT1 Then
    
        Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
        
        Temp1 = Temp1 * TSlope - Toffset
        
    End If
        
    txtTemp1 = Format$(Temp1, "##0.00")
        
    If EnableT2 Then
    
        Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
        
        Temp2 = Temp2 * TSlope - Toffset
    
    End If
    
    txtTemp2 = Format$(Temp2, "##0.00")
    
    'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
    If Not ValidSensorTemp(Temp1, Temp2) Then
    
        'Start code to tell user that the temp sensor values are bad
        NotifySensorError Temp1, Temp2
        
    End If
        
    
    If EnableT1 Or EnableT2 Then
        
        Do While Temp1 >= Thot Or Temp2 >= Thot
            
            frmADWIN_AF.ZOrder
            frmADWIN_AF.Show
            
            lblAFtooHot.Visible = True
            txtTemp1.BackColor = ColorOrange
            txtTemp2.BackColor = ColorOrange
            
            ErrorMessage = "The AF degaussing unit is above " & Thot & "°C: " & Format$(Temp1, "##0.00") & _
                "°C and " & Format$(Temp2, "##0.00") & "°C." & _
                vbCrLf & "Execution will restart soon."
            
            If TWarning = False Then frmSendMail.MailNotification "AF too hot", ErrorMessage, CodeYellow
            
            TWarning = True
            
            ' MsgBox "Pause... " & Temp1 & "°C " & Temp2 & "°C"
            ' Loop until the temperature which was above Thot decreases at least 5 degrees before restarting
            Do While Temp1 >= Thot - 5 Or Temp2 >= Thot - 5
                
                DelayTime (1)
                
                '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                If EnableT1 Then
                
                    Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
                    
                    Temp1 = Temp1 * TSlope - Toffset
                    
                End If
                
                txtTemp1 = Format$(Temp1, "##0.00")
                
                If EnableT2 Then
                
                    Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
                    
                    Temp2 = Temp2 * TSlope - Toffset
                
                End If
                
                txtTemp2 = Format$(Temp2, "##0.00")
            
                'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
                If Not ValidSensorTemp(Temp1, Temp2) Then
                
                    'Start code to tell user that the temp sensor values are bad
                    NotifySensorError Temp1, Temp2
                    
                End If
            
            Loop
        
        Loop
    
    End If
    
    txtTemp1.BackColor = RGB(255, 255, 255)
    txtTemp2.BackColor = RGB(255, 255, 255)
    
    lblAFtooHot.Visible = False
   
oops:
        
    '(April 30, 2011) I Hilburn
    'Added in check to see if the program flow is paused.  If so, the AF code will wait until the flow
    'is returned to Resume
    If Prog_paused = True Then
    
        'Set TempS equal to the current contents
        'of the 2nd status bar panel
        TempS = frmProgram.sbStatusBar.Panels(2).text
        
        'Update the program form status bar
        frmProgram.StatusBar "Paused...", 2
        
        'Loop and wait for status to change and DoEvents
        'to allow user to make changes
        Do
        
            'Pause 100 ms
            PauseTill timeGetTime() + 100
            
            DoEvents
            
        Loop Until Prog_paused = False
        
        'Return the old string to panel 2 of the status bar
        frmProgram.StatusBar TempS, 2
        
    End If
        
'------------------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------------------'

    'Turn off the error-pop-up in the Boot process
    ADWIN.Show_Errors (0)

    'Boot the ADWIN board if it isn't already
    If ADWIN.ADWIN_BootBoard(MonitorWave.BoardUsed) = False Then
    
        'Test to see if the Boot was successful
        DoRampADWIN = -1
        
        Exit Function
        
    End If
    
    'Now clear all the process on the ADWIN board
    ReturnVal = ADWIN.ClearAll_Processes
    
    'Otherwise, the board is booted and ready to go
        
    'Save the Ramp Process Name
    ProcessName = "AF Ramp"
        
    'Load the 1st process, the AF Ramp output
    ReturnVal = ADWIN.Load_Process(ADWIN.BinFolderPath & ADWIN.CurProcessFile)
    
    'ReturnVal of 1 = OK, ReturnVal <> 1 = Error occurred
    If ReturnVal <> 1 Then
    
        Flow_Pause
        SetCodeLevel CodeRed
        
        ErrorMessage = "Unable to load " & ProcessName & " process into the ADWIN board." & _
                       vbNewLine & vbNewLine & _
                       "Process File = " & ADWIN.CurProcessFile & vbNewLine & _
                       "ADWIN Dev No. = " & Trim(str(ADWIN.GetDeviceNo)) & vbNewLine & _
                       "Board Name = " & MonitorWave.BoardUsed.BoardName
    
        frmSendMail.MailNotification "ADWIN AF Error", _
                                     ErrorMessage, _
                                     CodeRed, _
                                     True
                                     
        ErrorMessage = "Unable to load " & ProcessName & " process into the ADWIN board." & _
                       vbNewLine & vbNewLine & _
                       "Process File = " & ADWIN.CurProcessFile & ", " & _
                       "ADWIN Dev No. = " & Trim(str(ADWIN.GetDeviceNo)) & ", " & _
                       "Board Name = " & MonitorWave.BoardUsed.BoardName & vbNewLine & vbNewLine & _
                       "Would you like to access the AF file settings right now to fix this? " & _
                       "If you click 'Cancel', the AF ramp will be aborted." & _
                       "If you click 'No', the AF ramp will continue."
                                     
        'Prompt the user to see if they want to change the ADWIN AF file settings
        UserResp = MsgBox(ErrorMessage, vbYesNoCancel, "ADWIN AF Settings Error")
        
        'Return code status to prior value
        frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
                
        If UserResp = vbYes Then
                
            'Load & show frmFileSave
            Load frmFileSave
            frmFileSave.Show
                            
            'Pop-up message box to remind user to set flow back to running
            'when done changing the settings
            MsgBox "Remember to set the program flow back to ""Running"" " & _
                   "in the main window 'Flow' menu or in the Settings form " & _
                   "after you are done editing the ADWIN AF file settings."
                
            'Call the ADWIN ramp function a second time.
            DoRampADWIN = DoRampADWIN(MonitorWave, _
                                      UpWave, _
                                      DownWave, _
                                      AF_Data(), _
                                      PeakField, _
                                      HangeTime, _
                                      RampMode)
                  
            'Exit the function to prevent infinite recursive error loops
            Exit Function
            
        ElseIf UserResp = vbCancel Then
        
            DoRampADWIN = -1
            
            Exit Function
            
        End If
        
    End If
    
'----------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------'
    
    'Now need to load all of the parameters for the processes
    
'----------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------'
'
'   Load Float Parameters first - must be passed in as Single data-type
'
'----------------------------------------------------------------------------------------------'
'
'   #define SLOPE_UP    FPAR_31 ' volts / second
'   #define SLOPE_DOWN  FPAR_32 ' volts / second
'   #define PEAKVOLTAGE FPAR_33 ' volts
'   #define FREQ       FPAR_34 ' frequency of DAC sine wave field in Hz
'   #define AC_AMPL_LIMIT   FPAR_35
'   #define MAX_RAMPVOLTAGE FPAR_36 ' volts - absolute max that ramp voltage can go up to
''                                    before the ramp is terminated
'   #define MAX_PEAKVOLTAGE FPAR_37 ' volts - absolute max that the input peak voltage can go up
''                                    to before the ramp is terminated
'----------------------------------------------------------------------------------------------'

    'Set the Slope Up and the Slope Down for the ramp cycle
    ADWIN.Set_Fpar 31, CSng(UpWave.Slope)
    ADWIN.Set_Fpar 32, CSng(DownWave.Slope)
    
    'Set the peak monitor voltage
    ADWIN.Set_Fpar 33, CSng(MonitorWave.PeakVoltage)
    
    'Set the Min & sine Wave freq,
    ADWIN.Set_Fpar 34, CSng(MonitorWave.SineFreqMin)
    
    'Set the Ramp Up suggested voltage peak limit
    ADWIN.Set_Fpar 35, CSng(UpWave.PeakVoltage)
    
    'Set the Absolute Max Ramp Output voltage
    If ActiveCoilSystem = AxialCoilSystem Then
    
        If modConfig.AfAxialRampMax = -1 Then
    
            ADWIN.Set_Fpar 36, CSng(10)
    
        Else
    
            ADWIN.Set_Fpar 36, CSng(modConfig.AfAxialRampMax)
            
        End If
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        If modConfig.AfTransRampMax = -1 Then
        
            ADWIN.Set_Fpar 36, CSng(10)
            
        Else
        
            ADWIN.Set_Fpar 36, CSng(modConfig.AfTransRampMax)
            
        End If
        
    End If
    
    'Set the absolute max monitor input voltage to accept
    If ActiveCoilSystem = AxialCoilSystem Then
    
        If modConfig.AfAxialMonMax = -1 Then
        
            ADWIN.Set_Fpar 37, CSng(10)
            
        Else
        
            ADWIN.Set_Fpar 37, CSng(modConfig.AfAxialMonMax)
            
        End If
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        If modConfig.AfTransMonMax = -1 Then
    
            ADWIN.Set_Fpar 37, CSng(10)
    
        Else
    
            ADWIN.Set_Fpar 37, CSng(modConfig.AfTransMonMax)
            
        End If
        
    End If
       
'----------------------------------------------------------------------------------------------'
'
'   Load the Long-type parameters next
'
'----------------------------------------------------------------------------------------------'
'
'    '--------------------------------------------------------------------------------
'    'Modifications for different styles of Ramp cycle needed by the Visual Basic
'    'GUI to do:
'    ' 1 - Normal, Monitored AF RAMP, without external storage of Ramp data
'    ' 2 - Debug mode, Monitored AF RAMP, with external storage of Ramp data
'    '     this mode will also be used to do Field Calibration of the AF Ramp in tandem
'    '     with a computer controlled Hirst 908A Guassmeter & Axial or Transverse Hall-probes
'    ' 3 - Clipping test, unmonitored AF RAMP, with external storage of Ramp data
'    ' 4 - AF Tuning run, unmonitored AF Ramp, with freq varying from min to max freq with
'    '     external storage of data
'    '--------------------------------------------------------------------------------
'    #define RAMPMODE PAR_31
'
'    '--------------------------------------------------------------------------------
'    'Modified values of three ports to be setable by passing in a parameter
'    'from the external GUI program.
'    '--------------------------------------------------------------------------------
'    #define PORT_SINEOUT PAR_32  '(March 2010, I Hilburn - creating global const to store the DAC-OUT port number for the sine wave output)
'    #define PORT_ACCUR PAR_33
'
'    '--------------------------------------------------------------------------------
'    'Speed in process cycle delay at which to run the AF Ramp IO process
'    '--------------------------------------------------------------------------------
'    #define AFRAMP_PD PAR_34
'
'    '--------------------------------------------------------------------------------
'    'Distance in 16-bit integer counts that max input value gets within the target
'    'peak before the ramp up process finishes
'    '--------------------------------------------------------------------------------
'    #define NOISELEVEL PAR_35
'    '--------------------------------------------------------------------------------
'    'Peak Delay time in Periods / Cycles
'    '--------------------------------------------------------------------------------
'    #define PEAKDELAY_PERIODS PAR_36
'
'    '--------------------------------------------------------------------------------
'    'Variables to Set the number of periods / cycles for the Ramp Down
'    'and to select between using the Ramp Down slope or using the number of periods
'    '
'    'RAMPDOWN_MODE = 0 -- use slope (Volts / second) to ramp down
'    'RAMPDONW_MODE = 1 -- use number of periods to ramp down
'    '--------------------------------------------------------------------------------
'    #define NUMPERIODS PAR_37
'    #define RAMPDOWN_MODE PAR_38
'----------------------------------------------------------------------------------------------'
    
    'Set the RampMode Parameter
    ADWIN.Set_Par 31, RampMode
                  
    'Set the AF Ramp DAC output port parameter
    ADWIN.Set_Par 32, UpWave.Chan.ChanNum
    
    'Set the AF monitor ADC input port parameter
    ADWIN.Set_Par 33, MonitorWave.Chan.ChanNum
    
'    Dim fso As FileSystemObject
'    Set fso = New FileSystemObject
'
'    Dim txt_stream As TextStream
'
'
'    If fso.FileExists("C:\Test.txt") Then
'
'        Set txt_stream = fso.OpenTextFile("C:\Test.txt", ForAppending)
'
'    Else
'        Set txt_stream = fso.CreateTextFile("C:\Test.txt")
'
'    End If
'
'    Dim debug_txt As String
'
'    debug_txt = "AF Target Voltage = " & Trim(Str(MonitorWave.PeakVoltage)) & ", " & _
'                "Sine Freq = " & Trim(Str(MonitorWave.SineFreqMin)) & ", " & _
'                "Target Ramp Voltage = " & Trim(Str(UpWave.PeakVoltage)) & ", " & _
'                "RampMode = " & Trim(Str(RampMode)) & ", " & _
'                "DAC chan = " & Trim(Str(UpWave.Chan.ChanNum)) & ", " & _
'                "ADC chan = " & Trim(Str(MonitorWave.Chan.ChanNum))
'
'    txt_stream.Write debug_txt & vbCrLf
'    txt_stream.Close
'    Set txt_stream = Nothing
'    Set fso = Nothing




   
    '??? /4/3/17
    'Set Slope = 1 to prevent crashes
    
    If MonitorWave.Slope = 0 Then
        MonitorWave.Slope = 1
    End If
    
    If UpWave.Slope = 0 Then
        UpWave.Slope = 1
    End If
    
    If DownWave.Slope = 0 Then
        DownWave.Slope = 1
    End If
    
 
   '??? 4/3/17
    
    'Set the AF ramp rate - with conversion of Hz to ADWIN Process delay with a 25 ns processor
    'cycling time
    ADWIN.Set_Par 34, CLng(1000000# / UpWave.IORate * 40)
    
    'Set the AF Noise level
    ADWIN.Set_Par 35, NoiseLevel
    
    'Set the number of periods to hang at peak
    '(1 / 1000 to convert miliseconds to seconds)
    ADWIN.Set_Par 36, CLng(HangeTime * MonitorWave.SineFreqMin / 1000)
    
    'Set the number of periods to ramp down with
    ADWIN.Set_Par 37, CLng(DownWave.PeakVoltage * MonitorWave.SineFreqMin / DownWave.Slope)
    
    'Set the Ramp-down mode:
    ' 0 = use the Ramp down slope
    ' 1 = use the Ramp down number of periods
    ADWIN.Set_Par 38, RampDownMode
    
    'All the necessary parameters have been set now
    
'----------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------'

    'Change the Status Bar in the program window to indicate that the ramp is being executed
    'First, put together the status string
    If Me.optCalRamp(0).value = True And Me.chkClippingTest.value = Unchecked Then
    
        'This is a calibrated ramp
        TempS = "AF Ramp: " & Trim(str(PeakField)) & " " & _
                Trim(Me.cmbFieldUnits.List(Me.cmbFieldUnits.ListIndex))
        
    ElseIf Me.chkClippingTest.value = Unchecked Then
    
        'This is an uncalibrated ramp
        TempS = "AF Ramp: " & Trim(str(MonitorWave.PeakVoltage)) & " V"
        
    Else
    
        'This is an unmonitored ramp
        TempS = "AF Ramp: " & Trim(str(UpWave.PeakVoltage)) & " V"
        
    End If

    'Update the status bar
    frmProgram.StatusBar TempS, 3
    
    'Calculate the Up & Down wave durations
    UpWave.Duration = CLng(UpWave.PeakVoltage / UpWave.Slope * 1000)
    DownWave.Duration = CLng(DownWave.PeakVoltage / DownWave.Slope * 1000)
    
    'Set Monitor Wave Duration
    MonitorWave.Duration = UpWave.Duration + _
                           DownWave.Duration + _
                           HangeTime + 200
    
    
    'Has the user selected to record the DC field?
    If DoDCFieldRecord = True Then
    
        'Set the DC Field Status to false
        DCFieldStatus = False
    
        'Set the DCFIeld Wave to a new wave object
        Set DCFieldWave = New Wave
    
        'Initialize the DC field record
        'For duration - put in the estimated duration of the ramp cycle
        'plus an additional 200 miliseconds
        frm908AGaussmeter.InitializeDCFieldRecord _
                                DCFieldWave, _
                                MonitorWave.Duration + 200
                                           
        'Start the DC Field record and grab the
        'status of the start process
        DCFieldStatus = frm908AGaussmeter.StartDCFieldRecord(DCFieldWave)
        
        Debug.Print Trim(str(DCFieldStatus))
        
    End If
    
    'Start the Ramp process
    ADWIN.Start_Process 1
    
    'Store the start time of the ramp
    StartTime = timeGetTime()
    
    'Pause 100 ms
    PauseTill timeGetTime() + 100
    
    'Loop while checking the ramp process on the ADWIN board every 200 milliseconds
    Do
        
        'Pause 200 ms
        PauseTill timeGetTime() + 200
    
        TempL = ADWIN.Get_Par(4)
    
    Loop Until TempL = 7
        
    'Get the elapsed time of the ramp
    RampDuration = timeGetTime() - StartTime
    
    'Stop the DC Field record if the user has set for it to be recorded
    'and the record was started successfully
    If DCFieldStatus = True And _
       DoDCFieldRecord = True _
    Then
    
        'Stop the DC Field record
        frm908AGaussmeter.StopDCFieldRecord DCFieldWave, True
        
    End If
           
    'Pause an additional 200 ms
    PauseTill timeGetTime() + 200
            
    'Now Get the basic parameters that we need to know about the AF Ramp
    'whether or not we're in debug or clipping mode
    
   'Get the final point of the DAC output (OUTCOUNT)
    DownWave.CurrentPoint = ADWIN.Get_Par(5)
    
    'Get the point # of the last ADC input point from the INCOUNT parameter
    MonitorWave.CurrentPoint = ADWIN.Get_Par(6)
    
    'Clip off 10 points from the end of the data set
    MonitorWave.CurrentPoint = MonitorWave.CurrentPoint - 10

    'Get the point # of the last DAC output point from the Ramp Up process
    UpWave.CurrentPoint = ADWIN.Get_Par(7)
    
    'Get the first point # of the DAC output Ramp Down process
    DownWave.StartPoint = ADWIN.Get_Par(8)
    
    'Get the max input voltage reached
    MonitorWave.CurrentVoltage = ADWIN.Get_Fpar(4)
    
    'Get the max output voltage reached
    UpWave.CurrentVoltage = ADWIN.Get_Fpar(5)
    
    'Get the Down slope (may now be different depending on the ramp mode)
    DownWave.Slope = ADWIN.Get_Fpar(32)
    
    'Get the Ramp process Time steps from the ADWIN board - ACOUT_TIMESTEP
    UpWave.TimeStep = ADWIN.Get_Fpar(6)
    MonitorWave.TimeStep = UpWave.TimeStep
    
    'Get the actual used points per period from the ADWIN board - NSAMPLES
    MonitorWave.PtsPerPeriod = ADWIN.Get_Fpar(7)
    
    'Quick error check
    'If the monitor max voltage is < 4 * Noiselevel * 20 / 2^16
    ' and the Ramp Voltage is at or above the suggested peak, then
    'the monitor channel settings are wrong or somehow messed up.
    With MonitorWave
    
        If (.CurrentVoltage < 4 * NoiseLevel * _
                              (.range.MaxValue - .range.MinValue) / 2 ^ 16) And _
           UpWave.CurrentVoltage >= UpWave.PeakVoltage _
        Then
        
            Flow_Pause
            SetCodeLevel CodeRed
        
            ErrorMessage = "Abnormally low monitor input voltage on AF " & .BoardUsed.BoardName & _
                           " board." & vbNewLine & _
                           "Target Monitor Voltage: " & Format(.PeakVoltage, "#0.000") & vbNewLine & _
                           "Current Monitor Voltage: " & Format(.CurrentVoltage, "#0.000") & _
                           vbNewLine & vbNewLine & _
                           "Code execution has been paused.  Please come and check the machine."
        
            'Something's wrong with the ADWIN board input voltage channel
            'Raise a code red error and send an email
            frmSendMail.MailNotification _
                        "AF Monitor Error.", _
                        ErrorMessage, _
                        CodeRed, _
                        True
                        
            ErrorMessage = "Abnormally low monitor input voltage on AF " & .BoardUsed.BoardName & _
                           " board.  Ramp voltage output and/or monitor voltage input comm settings " & _
                           "may be wrong." & vbNewLine & vbNewLine & _
                           "Would you like to change the ADWIN AF comm settings?" & vbNewLine & _
                           "Cancel = abort AF ramp" & vbNewLine & _
                           "No = continue Paleomag code without changing the settings."
            
            UserResp = MsgBox(ErrorMessage, _
                              vbYesNoCancel, _
                              "ADWIN Comm Error?")
                              
            frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
            
            If UserResp = vbYes Then
            
                'Load & show frmSettings and frmADWIN_AF_CommSettings
                Load frmSettings
                frmSettings.ZOrder
                frmSettings.selectTab 4
                frmSettings.Show
                    
                'Pop-up message box to remind user to set flow back to running
                'when done changing the settings
                MsgBox "Remember to set the program flow back to ""Running"" " & _
                       "in the main window 'Flow' menu or in the Settings form " & _
                       "after you are done editing the ADWIN AF comm. settings."
            
                'Restart the ADWIN ramp
                DoRampADWIN = DoRampADWIN(MonitorWave, _
                                          UpWave, _
                                          DownWave, _
                                          AF_Data(), _
                                          PeakField, _
                                          HangeTime, _
                                          RampMode, _
                                          DoDCFieldRecord)
            
                'Exit function to prevent infinite recursive error loops!
                Exit Function
            
            ElseIf UserResp = vbCancel Then
            
                'Blank the status bar
                frmProgram.StatusBar vbNullString, 3
            
                'Abort the AF ramp
                DoRampADWIN = -1
                
                Exit Function
                
            End If
            
        End If
        
    End With
            
    'Now need to retrieve the Ramp Data, if Verbose = True
    If RampMode > 1 Then
    
        'Update Status bar again
        frmProgram.StatusBar "Getting data...", 3
    
        'Pause for 1/4 the duration of the last ramp cycle for Ramp Data arrays
        'to become available
        PauseTill timeGetTime() + RampDuration \ 4
        
        'Set N = the maximum # of data points that can be stored by
        'the ADWIN code = MAXALLOWEDDATAPTS
        N = ADWIN.Get_Par(11)

'---------------------------------------------------------------------------------------------'
'
'   April 29, 2010
'   I Hilburn
'
'   This code is now obsolete, could cause some problems downstream if this function
'   is called during an AF Tune, Auto-clip test, or Auto-calibrate run
'
'---------------------------------------------------------------------------------------------'
'        'If the INCOUNT or RAmp up count, or Ramp Down start count are
'        'larger than N, then then these values need to be altered
'        'No data has been stored by the ADWIN board for ramp points with INCOUNT > N
'        If N < MONITORWAVE.CurrentPoint Then MONITORWAVE.CurrentPoint = N
'        If N < UpWave.CurrentPoint Then
'
'            UpWave.CurrentPoint = N
'            DownWave.StartPoint = -1
'
'        ElseIf N < DownWave.StartPoint Then
'
'            DownWave.StartPoint = -1
'
'        End If
'---------------------------------------------------------------------------------------------'
                
        'Redimension the Temp & AF_Data arrays so that they are the
        'Same size as the number of INCOUNT points
        
        'If this is a calibrated ramp, need three columns in AF_Data
        If optCalRamp(0).value = True Then
        
            ReDim AF_Data(MonitorWave.CurrentPoint, 3)
            
        Else
        
            ReDim AF_Data(MonitorWave.CurrentPoint, 2)
            
        End If
        
        ReDim TempDataIn(MonitorWave.CurrentPoint + 1)
        ReDim TempDataOut(MonitorWave.CurrentPoint + 1)
        
        'No get the data from the LONG valued ADWIN external memory data arrays
        'loaded into TempDataIn and TempDataOut
        ADWIN.GetData_Long 31, 1, MonitorWave.CurrentPoint, TempDataIn
        ADWIN.GetData_Long 32, 1, MonitorWave.CurrentPoint, TempDataOut
        
        'Set Percent done to 0%
        PercentDone = "  0%"
        
        'Data has been retrieved, change status bar status to "Converting..."
        frmProgram.StatusBar "Converting... " & PercentDone, 3
        
        
        'If the user has run a calibrated ramp, need to store an additional column
        'containing the Ramp Field value. This value is just the Monitor voltage
        'rescaled by the ratio of the Peak Field by the Peak Monitor Voltage value
        If optCalRamp(0).value = True Then
        
            'Calculate the ratio of the Peak Field to the Peak Monitor Voltage
            TempD = PeakField / MonitorWave.PeakVoltage
        
            For i = 1 To MonitorWave.CurrentPoint
            
                AF_Data(i - 1, 0) = UpWave.range.ADWIN_RangeConverter(, TempDataIn(i))
                AF_Data(i - 1, 1) = UpWave.range.ADWIN_RangeConverter(, TempDataOut(i))
                AF_Data(i - 1, 2) = TempD * _
                                    UpWave.range.ADWIN_RangeConverter(, TempDataIn(i))
                                    
                'Every 1 hundred points, update data-conversion status
                If i Mod 100 = 0 Then
                
                    'Format the percent done string
                    PercentDone = Trim(str(CInt(i / MonitorWave.CurrentPoint * 100)))
                    PercentDone = PadLeft(PercentDone, 4) & "%"
                    
                    'Update the program form status bar
                    frmProgram.StatusBar "Converting... " & PercentDone, 3
                    
                End If
                
            Next i
            
        Else
            
            For i = 1 To MonitorWave.CurrentPoint
            
                AF_Data(i - 1, 0) = UpWave.range.ADWIN_RangeConverter(, TempDataIn(i))
                AF_Data(i - 1, 1) = UpWave.range.ADWIN_RangeConverter(, TempDataOut(i))
                
                'Every 1 hundred points, update data-conversion status
                If i Mod 100 = 0 Then
                
                    'Format the percent done string
                    PercentDone = Trim(str(CInt(i / MonitorWave.CurrentPoint * 100)))
                    PercentDone = PadLeft(PercentDone, 4) & "%"
                    
                    'Update the program form status bar
                    frmProgram.StatusBar "Converting... " & PercentDone, 3
                    
                End If
                
            Next i
            
        End If
                                           
    End If
        
    'Clear all processes on the ADWIN board
    ReturnVal = ADWIN.ClearAll_Processes
        
    'Reset the Program form status bar to null string
    frmProgram.StatusBar vbNullString, 3
        
    'Deallocate the DC field record wave
    Set DCFieldWave = Nothing
        
    DoRampADWIN = 0
    
End Function

Public Function DoRampADWIN_WithParameterLogging( _
    ByRef MonitorWave As Wave, _
    ByRef UpWave As Wave, _
    ByRef DownWave As Wave, _
    ByRef AF_Data() As Double, _
    ByVal PeakField As Double, _
    Optional ByVal HangeTime As Long = 0, _
    Optional ByVal RampMode As Long = 1, _
    Optional ByVal RampDownMode As Long = 0, _
    Optional ByVal DoDCFieldRecord As Boolean = False, _
    Optional ByVal RetryNumber As Integer = 0) As Long
                       
    Dim ReturnVal As Long
    Dim ProcessFile As String
    Dim ProcessName As String
    Dim PercentDone As String
    
    Dim i As Long
    Dim N As Long
    Dim TempDataIn() As Long
    Dim TempDataOut() As Long
    
    Dim TempD As Double
    Dim TempS As String
    Dim TempL As Long
    Dim Temp1 As Double
    Dim Temp2 As Double
    Dim TWarning As Boolean
    
    Dim StartTime As Long
    Dim RampDuration As Long
    Dim UserResp As Long
    Dim NumPeriods As Long
        
    Dim ErrorMessage As String
    
    Dim adwin_af_error As af_ramp_error
    Set adwin_af_error = New af_ramp_error
    
    Dim DCFieldWave As Wave
    Dim DCFieldStatus As Boolean
    
    'Exit if nocomm mode
    If NOCOMM_MODE = True Then
    
        'This is not an error free ramp
        DoRampADWIN_WithParameterLogging = -616
        
        Exit Function
        
    End If
    
    '(July 23, 2010 - I Hilburn)
    'Added in check to see if the program flow is paused.  If so, the AF code will wait until the flow
    'is returned to Resume
    If Prog_paused = True Then
    
        'Set TempS equal to the current contents
        'of the 2nd status bar panel
        TempS = frmProgram.sbStatusBar.Panels(2).text
        
        'Update the program form status bar
        frmProgram.StatusBar "Paused...", 2
        
        'Loop and wait for status to change and DoEvents
        'to allow user to make changes
        Do
        
            'Pause 100 ms
            PauseTill timeGetTime() + 100
            
            DoEvents
            
        Loop Until Prog_paused = False
        
        'Return the old string to panel 2 of the status bar
        frmProgram.StatusBar TempS, 2
        
    End If
 
        
    'Do a double check real quick of the max ramp & monitor voltages
    '???
    If (ActiveCoilSystem = AxialCoilSystem And _
        (modConfig.AfAxialRampMax = -1 Or _
         modConfig.AfAxialMonMax = -1) And _
        RampMode <> 3) Or _
       (ActiveCoilSystem = TransverseCoilSystem And _
        (modConfig.AfTransRampMax = -1 Or _
         modConfig.AfTransMonMax = -1) And _
        RampMode <> 3) Or _
        (ActiveCoilSystem = IRMAxialCoilSystem And _
            (modConfig.AfAxialRampMax = -1 Or _
             modConfig.AfAxialMonMax = -1) And _
            RampMode <> 3) _
        Then
        'Send an email and pop-up an error message
        SetCodeLevel CodeRed
        
        'Pause the flow
        Flow_Pause
        
        adwin_af_error.CoilType = GetCoilType_FromActiveAFAystem(ActiveCoilSystem)
        adwin_af_error.ErrorType = FatalError
        adwin_af_error.Source = "frmADWIN_AF.DoRampADWIN_WithParameterLogging"
        adwin_af_error.TimeStamp = Now
        
        ErrorMessage = "The Maximum Ramp voltages allowed for the ADWIN AF system " & _
                       "have not been set yet.  Please perform an AF Clipping-Test now."

        adwin_af_error.Message = ErrorMessage
        
        modAF_DAQ.UpdateAFErrorStatus adwin_af_error

        'Send the error notification
        frmSendMail.MailNotification "AF Settings Error", _
                                     ErrorMessage, _
                                     CodeRed, _
                                     True
                                     
        MsgBox "The Maximum Ramp voltages allowed for the ADWIN AF system " & _
               "have not been set yet.  Please perform an AF Clipping-Test now.", _
               vbOK, _
               "AF Settings Error"
    
        Load frmAFTuner
        frmAFTuner.Show
        
        Me.Hide
        
        DoRampADWIN_WithParameterLogging = -616
        
        Exit Function
        
    End If
        
    'Update status bar to show that the AF run is being configured
    frmProgram.StatusBar "AF Config", 3
        
'------------------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------------------'
'
'   July 2010
'   Authors:
'   Laurent Corporozen
'   Isaac Hilburn
'
'   Copied temperature check code from frmAF_2G.ExecuteRamp to this ADWIN central ramp function
'   Code copied verbatim with minor changes to switch comm implementation from frmMCC to frmDAQ_Comm
'   frmDAQ_Comm takes Channel objects instead of channel port numbers
'------------------------------------------------------------------------------------------------------------------------------------'
        
    'Before doing anything with the ADWIN board, get the AF coil temperatures
    '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
    If EnableT1 Then
    
        Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
        
        Temp1 = Temp1 * TSlope - Toffset
        
    End If
        
    txtTemp1 = Format$(Temp1, "##0.00")
        
    If EnableT2 Then
    
        Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
        
        Temp2 = Temp2 * TSlope - Toffset
    
    End If
    
    txtTemp2 = Format$(Temp2, "##0.00")
    
    'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
    If Not ValidSensorTemp(Temp1, Temp2) Then
    
        'Start code to tell user that the temp sensor values are bad
        NotifySensorError Temp1, Temp2
        
    End If
        
    
    If EnableT1 Or EnableT2 Then
        
        Do While Temp1 >= Thot Or Temp2 >= Thot
            
            frmADWIN_AF.ZOrder
            frmADWIN_AF.Show
            
            lblAFtooHot.Visible = True
            txtTemp1.BackColor = ColorOrange
            txtTemp2.BackColor = ColorOrange
            
            ErrorMessage = "The AF degaussing unit is above " & Thot & "°C: " & Format$(Temp1, "##0.00") & _
                "°C and " & Format$(Temp2, "##0.00") & "°C." & _
                vbCrLf & "Execution will restart soon."
            
            If TWarning = False Then frmSendMail.MailNotification "AF too hot", ErrorMessage, CodeYellow
            
            TWarning = True
            
            ' MsgBox "Pause... " & Temp1 & "°C " & Temp2 & "°C"
            ' Loop until the temperature which was above Thot decreases at least 5 degrees before restarting
            Do While Temp1 >= Thot - 5 Or Temp2 >= Thot - 5
                
                DelayTime (1)
                
                '(July 2010 - I Hilburn) Replaces old frmMCC with frmDAQ_Comm using the Channel object variables
                If EnableT1 Then
                
                    Temp1 = frmDAQ_Comm.DoDAQIO(AnalogT1)
                    
                    Temp1 = Temp1 * TSlope - Toffset
                    
                End If
                
                txtTemp1 = Format$(Temp1, "##0.00")
                
                If EnableT2 Then
                
                    Temp2 = frmDAQ_Comm.DoDAQIO(AnalogT2)
                    
                    Temp2 = Temp2 * TSlope - Toffset
                
                End If
                
                txtTemp2 = Format$(Temp2, "##0.00")
            
                'Check Temperature to see if it is not zeroed (gone within 20 deg of -1 * Toffset)
                If Not ValidSensorTemp(Temp1, Temp2) Then
                
                    'Start code to tell user that the temp sensor values are bad
                    NotifySensorError Temp1, Temp2
                    
                End If
            
            Loop
        
        Loop
    
    End If
    
    txtTemp1.BackColor = RGB(255, 255, 255)
    txtTemp2.BackColor = RGB(255, 255, 255)
    
    lblAFtooHot.Visible = False
   
oops:
        
    '(April 30, 2011) I Hilburn
    'Added in check to see if the program flow is paused.  If so, the AF code will wait until the flow
    'is returned to Resume
    If Prog_paused = True Then
    
        'Set TempS equal to the current contents
        'of the 2nd status bar panel
        TempS = frmProgram.sbStatusBar.Panels(2).text
        
        'Update the program form status bar
        frmProgram.StatusBar "Paused...", 2
        
        'Loop and wait for status to change and DoEvents
        'to allow user to make changes
        Do
        
            'Pause 100 ms
            PauseTill timeGetTime() + 100
            
            DoEvents
            
        Loop Until Prog_paused = False
        
        'Return the old string to panel 2 of the status bar
        frmProgram.StatusBar TempS, 2
        
    End If
        
'------------------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------------------'

    'Turn off the error-pop-up in the Boot process
    ADWIN.Show_Errors (0)

    'Boot the ADWIN board if it isn't already
    If ADWIN.ADWIN_BootBoard(MonitorWave.BoardUsed) = False Then
    
        'Test to see if the Boot was successful
        DoRampADWIN_WithParameterLogging = -1
        
        Exit Function
        
    End If
    
    'Now clear all the process on the ADWIN board
    ReturnVal = ADWIN.ClearAll_Processes
    
    
    Dim pause_constants As AdwinAfPauseConstants
    Set pause_constants = New AdwinAfPauseConstants
    
    PauseTill_NoEvents timeGetTime() + pause_constants.MsecsBetweenBootAndInit
        
    'Save the Ramp Process Name
    ProcessName = "AF Ramp"
        
    'Load the 1st process, the AF Ramp output
    ReturnVal = ADWIN.Load_Process(ADWIN.BinFolderPath & ADWIN.CurProcessFile)
    
    'ReturnVal of 1 = OK, ReturnVal <> 1 = Error occurred
    If ReturnVal <> 1 Then
    
        Flow_Pause
        SetCodeLevel CodeRed
        
        ErrorMessage = "Unable to load " & ProcessName & " process into the ADWIN board." & _
                       vbNewLine & vbNewLine & _
                       "Process File = " & ADWIN.CurProcessFile & vbNewLine & _
                       "ADWIN Dev No. = " & Trim(str(ADWIN.GetDeviceNo)) & vbNewLine & _
                       "Board Name = " & MonitorWave.BoardUsed.BoardName
    
        adwin_af_error.CoilType = GetCoilType_FromActiveAFAystem(ActiveCoilSystem)
        adwin_af_error.ErrorType = FatalError
        adwin_af_error.Source = "frmADWIN_AF.DoRampADWIN_WithParameterLogging"
        adwin_af_error.TimeStamp = Now
        
        adwin_af_error.Message = ErrorMessage
        
        modAF_DAQ.UpdateAFErrorStatus adwin_af_error
    
        frmSendMail.MailNotification "ADWIN AF Error", _
                                     ErrorMessage, _
                                     CodeRed, _
                                     True
                                     
        ErrorMessage = "Unable to load " & ProcessName & " process into the ADWIN board." & _
                       vbNewLine & vbNewLine & _
                       "Process File = " & ADWIN.CurProcessFile & ", " & _
                       "ADWIN Dev No. = " & Trim(str(ADWIN.GetDeviceNo)) & ", " & _
                       "Board Name = " & MonitorWave.BoardUsed.BoardName & vbNewLine & vbNewLine & _
                       "Would you like to access the AF file settings right now to fix this? " & _
                       "If you click 'Cancel', the AF ramp will be aborted." & _
                       "If you click 'No', the AF ramp will continue."
                                     
        'Prompt the user to see if they want to change the ADWIN AF file settings
        UserResp = MsgBox(ErrorMessage, vbYesNoCancel, "ADWIN AF Settings Error")
        
        'Return code status to prior value
        frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
                
        If UserResp = vbYes Then
                
            'Load & show frmFileSave
            Load frmFileSave
            frmFileSave.Show
                            
            'Pop-up message box to remind user to set flow back to running
            'when done changing the settings
            MsgBox "Remember to set the program flow back to ""Running"" " & _
                   "in the main window 'Flow' menu or in the Settings form " & _
                   "after you are done editing the ADWIN AF file settings."
                
            'Call the ADWIN ramp function a second time.
            DoRampADWIN_WithParameterLogging = DoRampADWIN_WithParameterLogging(MonitorWave, _
                                      UpWave, _
                                      DownWave, _
                                      AF_Data(), _
                                      PeakField, _
                                      HangeTime, _
                                      RampMode, _
                                      RampDownMode, _
                                      DoDCFieldRecord, _
                                      RetryNumber)
                  
            'Exit the function to prevent infinite recursive error loops
            Exit Function
            
        ElseIf UserResp = vbCancel Then
        
            DoRampADWIN_WithParameterLogging = -1
            
            Exit Function
            
        End If
        
    End If
    
'----------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------'
    
    'Now need to load all of the parameters for the processes
    
    Dim ramp_status As AdwinAfRampStatus
    Dim ramp_inputs As AdwinAfInputParameters
    Dim ramp_outputs As AdwinAfOutputParameters
    
    Set ramp_status = New AdwinAfRampStatus
    Set ramp_inputs = New AdwinAfInputParameters
    Set ramp_outputs = New AdwinAfOutputParameters
    
    If Me.optCalRamp.Item(0).value = True Then
       
       ramp_status.TargetPeakField = Trim(CStr(PeakField)) & " " & modConfig.AFUnits
       
    Else
    
       ramp_status.TargetPeakField = "uncal."
        
    End If
            
'----------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------'
'
'   Load Float Parameters first - must be passed in as Single data-type
'
'----------------------------------------------------------------------------------------------'
'
'   #define SLOPE_UP    FPAR_31 ' volts / second
'   #define SLOPE_DOWN  FPAR_32 ' volts / second
'   #define PEAKVOLTAGE FPAR_33 ' volts
'   #define FREQ       FPAR_34 ' frequency of DAC sine wave field in Hz
'   #define AC_AMPL_LIMIT   FPAR_35
'   #define MAX_RAMPVOLTAGE FPAR_36 ' volts - absolute max that ramp voltage can go up to
''                                    before the ramp is terminated
'   #define MAX_PEAKVOLTAGE FPAR_37 ' volts - absolute max that the input peak voltage can go up
''                                    to before the ramp is terminated
'----------------------------------------------------------------------------------------------'
    
    Dim CoilType As coil_type
    
    If ActiveCoilSystem = AxialCoilSystem Then
    
        ramp_inputs.Coil = "Axial"
        CoilType = coil_type.Axial
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        ramp_inputs.Coil = "Transverse"
        CoilType = coil_type.Transverse
    '??? //4/4/17
    ElseIf ActiveCoilSystem = IRMAxialCoilSystem Then
    
        ramp_inputs.Coil = "Axial"
        CoilType = coil_type.IRMAxial
    '??? 4/4/17
        
    Else
    
        ramp_inputs.Coil = "unknown"
        CoilType = coil_type.Unknown
        
    End If

    'Set the Slope Up and the Slope Down for the ramp cycle
    ADWIN.Set_Fpar 31, CSng(UpWave.Slope)
    ramp_inputs.Slope_Up.TrySetValue (CSng(UpWave.Slope))
    
    ADWIN.Set_Fpar 32, CSng(DownWave.Slope)
    ramp_inputs.Slope_Down.TrySetValue (CSng(DownWave.Slope))
    
    'Set the peak monitor voltage
    ADWIN.Set_Fpar 33, CSng(MonitorWave.PeakVoltage)
    ramp_inputs.Peak_Monitor_Voltage.TrySetValue (CSng(MonitorWave.PeakVoltage))
    
    'Set the Min & sine Wave freq,
    ADWIN.Set_Fpar 34, CSng(MonitorWave.SineFreqMin)
    ramp_inputs.Resonance_Freq.TrySetValue (CSng(MonitorWave.SineFreqMin))
    
    'Set the Ramp Up suggested ramp output voltage peak limit
    ADWIN.Set_Fpar 35, CSng(UpWave.PeakVoltage)
    ramp_inputs.Peak_Ramp_Voltage.TrySetValue (CSng(UpWave.PeakVoltage))
                    
    ramp_status.Coil = ""
                    
    'Set the Absolute Max Ramp Output voltage
    If ActiveCoilSystem = AxialCoilSystem Then
    
        ramp_status.Coil = "Axial"
        
        If modConfig.AfAxialRampMax = -1 Then
    
            ADWIN.Set_Fpar 36, CSng(10)
            ramp_inputs.Max_Ramp_Voltage.TrySetValue CSng(10)
    
        Else
    
            ADWIN.Set_Fpar 36, CSng(modConfig.AfAxialRampMax)
            ramp_inputs.Max_Ramp_Voltage.TrySetValue CSng(modConfig.AfAxialRampMax)
            
        End If
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        ramp_status.Coil = "Transverse"
        If modConfig.AfTransRampMax = -1 Then
        
            ADWIN.Set_Fpar 36, CSng(10)
            ramp_inputs.Max_Ramp_Voltage.TrySetValue CSng(10)
            
        Else
        
            ADWIN.Set_Fpar 36, CSng(modConfig.AfTransRampMax)
            ramp_inputs.Max_Ramp_Voltage.TrySetValue CSng(modConfig.AfTransRampMax)
            
        End If
        '??? //4/4/17
    ElseIf ActiveCoilSystem = IRMAxialCoilSystem Then
        ramp_status.Coil = "Axial"
        
        If modConfig.AfAxialRampMax = -1 Then
    
            ADWIN.Set_Fpar 36, CSng(10)
            ramp_inputs.Max_Ramp_Voltage.TrySetValue CSng(10)
    
        Else
    
            ADWIN.Set_Fpar 36, CSng(modConfig.AfAxialRampMax)
            ramp_inputs.Max_Ramp_Voltage.TrySetValue CSng(modConfig.AfAxialRampMax)
            
        End If
        '??? //4/4/17
    End If
    
    'Set the absolute max monitor input voltage to accept
    If ActiveCoilSystem = AxialCoilSystem Then
    
        If modConfig.AfAxialMonMax = -1 Then
        
            ADWIN.Set_Fpar 37, CSng(10)
            ramp_inputs.Max_Monitor_Voltage.TrySetValue CSng(10)
            
        Else
        
            ADWIN.Set_Fpar 37, CSng(modConfig.AfAxialMonMax)
            ramp_inputs.Max_Monitor_Voltage.TrySetValue CSng(modConfig.AfAxialMonMax)
            
        End If
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        If modConfig.AfTransMonMax = -1 Then
    
            ADWIN.Set_Fpar 37, CSng(10)
            ramp_inputs.Max_Monitor_Voltage.TrySetValue CSng(10)
    
        Else
    
            ADWIN.Set_Fpar 37, CSng(modConfig.AfTransMonMax)
            ramp_inputs.Max_Monitor_Voltage.TrySetValue CSng(modConfig.AfTransMonMax)
            
        End If
        '??? //4/4/17
    ElseIf ActiveCoilSystem = IRMAxialCoilSystem Then
    
        If modConfig.AfAxialMonMax = -1 Then
        
            ADWIN.Set_Fpar 37, CSng(10)
            ramp_inputs.Max_Monitor_Voltage.TrySetValue CSng(10)
            
        Else
        
            ADWIN.Set_Fpar 37, CSng(modConfig.AfAxialMonMax)
            ramp_inputs.Max_Monitor_Voltage.TrySetValue CSng(modConfig.AfAxialMonMax)
            
        End If

        '??? //4/4/17
    End If
       
'----------------------------------------------------------------------------------------------'
'
'   Load the Long-type parameters next
'
'----------------------------------------------------------------------------------------------'
'
'    '--------------------------------------------------------------------------------
'    'Modifications for different styles of Ramp cycle needed by the Visual Basic
'    'GUI to do:
'    ' 1 - Normal, Monitored AF RAMP, without external storage of Ramp data
'    ' 2 - Debug mode, Monitored AF RAMP, with external storage of Ramp data
'    '     this mode will also be used to do Field Calibration of the AF Ramp in tandem
'    '     with a computer controlled Hirst 908A Guassmeter & Axial or Transverse Hall-probes
'    ' 3 - Clipping test, unmonitored AF RAMP, with external storage of Ramp data
'    ' 4 - AF Tuning run, unmonitored AF Ramp, with freq varying from min to max freq with
'    '     external storage of data
'    '--------------------------------------------------------------------------------
'    #define RAMPMODE PAR_31
'
'    '--------------------------------------------------------------------------------
'    'Modified values of three ports to be setable by passing in a parameter
'    'from the external GUI program.
'    '--------------------------------------------------------------------------------
'    #define PORT_SINEOUT PAR_32  '(March 2010, I Hilburn - creating global const to store the DAC-OUT port number for the sine wave output)
'    #define PORT_ACCUR PAR_33
'
'    '--------------------------------------------------------------------------------
'    'Speed in process cycle delay at which to run the AF Ramp IO process
'    '--------------------------------------------------------------------------------
'    #define AFRAMP_PD PAR_34
'
'    '--------------------------------------------------------------------------------
'    'Distance in 16-bit integer counts that max input value gets within the target
'    'peak before the ramp up process finishes
'    '--------------------------------------------------------------------------------
'    #define NOISELEVEL PAR_35
'    '--------------------------------------------------------------------------------
'    'Peak Delay time in Periods / Cycles
'    '--------------------------------------------------------------------------------
'    #define PEAKDELAY_PERIODS PAR_36
'
'    '--------------------------------------------------------------------------------
'    'Variables to Set the number of periods / cycles for the Ramp Down
'    'and to select between using the Ramp Down slope or using the number of periods
'    '
'    'RAMPDOWN_MODE = 0 -- use slope (Volts / second) to ramp down
'    'RAMPDONW_MODE = 1 -- use number of periods to ramp down
'    '--------------------------------------------------------------------------------
'    #define NUMPERIODS PAR_37
'    #define RAMPDOWN_MODE PAR_38
'----------------------------------------------------------------------------------------------'

   
    '??? /4/3/17
    'Set Slope = 1 to prevent crashes
    
    If MonitorWave.Slope = 0 Then
        MonitorWave.Slope = 1
    End If
    
    If UpWave.Slope = 0 Then
        UpWave.Slope = 1
    End If
    
    If DownWave.Slope = 0 Then
        DownWave.Slope = 1
    End If
    
    If UpWave.IORate = 0 Then
        UpWave.IORate = 1
    End If
   '??? 4/3/17



    
    'Set the RampMode Parameter
    ADWIN.Set_Par 31, RampMode
    ramp_inputs.ramp_mode.TrySetValue (RampMode)
                  
    'Set the AF Ramp DAC output port parameter
    ADWIN.Set_Par 32, UpWave.Chan.ChanNum
    ramp_inputs.Output_Port_Number.TrySetValue (UpWave.Chan.ChanNum)
    
    'Set the AF monitor ADC input port parameter
    ADWIN.Set_Par 33, MonitorWave.Chan.ChanNum
    ramp_inputs.Monitor_Port_Number.TrySetValue (MonitorWave.Chan.ChanNum)
      
    'Set the AF ramp rate - with conversion of Hz to ADWIN Process delay with a 25 ns processor
    'cycling time
    ADWIN.Set_Par 34, CLng(1000000# / UpWave.IORate * 40)
    ramp_inputs.Process_Delay.TrySetValue (CLng(1000000# / UpWave.IORate * 40))
    
    'Set the AF Noise level
    ADWIN.Set_Par 35, NoiseLevel
    ramp_inputs.Noise_Level.TrySetValue (NoiseLevel)
    
    'Set the number of periods to hang at peak
    '(1 / 1000 to convert miliseconds to seconds)
    ADWIN.Set_Par 36, CLng(HangeTime * MonitorWave.SineFreqMin / 1000)
    ramp_inputs.Number_Periods_Hang_At_Peak.TrySetValue (CLng(HangeTime * MonitorWave.SineFreqMin / 1000))
    
    'Set the number of periods to ramp down with
    ADWIN.Set_Par 37, CLng(DownWave.PeakVoltage * MonitorWave.SineFreqMin / DownWave.Slope)
    ramp_inputs.Number_Periods_Ramp_Down.TrySetValue (CLng(DownWave.PeakVoltage * MonitorWave.SineFreqMin / DownWave.Slope))
    
    'Set the Ramp-down mode:
    ' 0 = use the Ramp down slope
    ' 1 = use the Ramp down number of periods
    ADWIN.Set_Par 38, RampDownMode
    ramp_inputs.ramp_down_mode.TrySetValue (RampDownMode)
    
    'All the necessary parameters have been set now
    
'----------------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------------'

    'Change the Status Bar in the program window to indicate that the ramp is being executed
    'First, put together the status string
    If Me.optCalRamp(0).value = True And Me.chkClippingTest.value = Unchecked Then
    
        'This is a calibrated ramp
        TempS = "AF Ramp: " & Trim(str(PeakField)) & " " & _
                Trim(Me.cmbFieldUnits.List(Me.cmbFieldUnits.ListIndex))
        
    ElseIf Me.chkClippingTest.value = Unchecked Then
    
        'This is an uncalibrated ramp
        TempS = "AF Ramp: " & Trim(str(MonitorWave.PeakVoltage)) & " V"
        
    Else
    
        'This is an unmonitored ramp
        TempS = "AF Ramp: " & Trim(str(UpWave.PeakVoltage)) & " V"
        
    End If

    'Update the status bar
    frmProgram.StatusBar TempS, 3
    
    'Calculate the Up & Down wave durations
    UpWave.Duration = CLng(UpWave.PeakVoltage / UpWave.Slope * 1000)
    DownWave.Duration = CLng(DownWave.PeakVoltage / DownWave.Slope * 1000)
    
    'Set Monitor Wave Duration
    MonitorWave.Duration = UpWave.Duration + _
                           DownWave.Duration + _
                           HangeTime + 200
    
    PauseTill_NoEvents timeGetTime() + pause_constants.MsecsBetweenInitAndRun
    
    'Has the user selected to record the DC field?
    If DoDCFieldRecord = True Then
    
        'Set the DC Field Status to false
        DCFieldStatus = False
    
        'Set the DCFIeld Wave to a new wave object
        Set DCFieldWave = New Wave
    
        'Initialize the DC field record
        'For duration - put in the estimated duration of the ramp cycle
        'plus an additional 200 miliseconds
        frm908AGaussmeter.InitializeDCFieldRecord _
                                DCFieldWave, _
                                MonitorWave.Duration + 200
                                           
        'Start the DC Field record and grab the
        'status of the start process
        DCFieldStatus = frm908AGaussmeter.StartDCFieldRecord(DCFieldWave)
        
        Debug.Print Trim(str(DCFieldStatus))
        
    End If
    
    'Verify the AF Relay position
    If Not IsCorrect_AFRelay_position(CoilType) Then
    
    '??? //4/4/17
    Dim returnFromTrySetRelays_ADwin As Boolean
    
    If (CoilType = IRMAxial) Then
        returnFromTrySetRelays_ADwin = frmADWIN_AF.TrySetRelays_ADwin(CoilType, True) 'True = use backfield
    Else
        returnFromTrySetRelays_ADwin = frmADWIN_AF.TrySetRelays_ADwin(CoilType)
    End If

    '??? //4/4/17
    
        'If Not frmADWIN_AF.TrySetRelays_ADwin(CoilType) Then '??? 4/4/17
        If Not returnFromTrySetRelays_ADwin Then '??? 4/4/17
        
            Flow_Pause
            SetCodeLevel CodeRed
            
            ErrorMessage = "Unable to set AF Relays to the correct position for ADwin AF " & ramp_status.Coil & " Ramp. " & _
                           vbNewLine & vbNewLine & _
                           "Code execution has been paused. AF Ramp has NOT yet been started. Please come and check the machine."
            
            adwin_af_error.CoilType = CoilType
            adwin_af_error.ErrorType = FatalError
            adwin_af_error.Source = "frmADWIN_AF.DoRampADWIN_WithParameterLogging"
            adwin_af_error.TimeStamp = Now
            
            adwin_af_error.Message = ErrorMessage
            
            modAF_DAQ.UpdateAFErrorStatus adwin_af_error
            
            'Something's wrong with the relays
            frmSendMail.MailNotification _
                        "AF Relay position error.", _
                        ErrorMessage, _
                        CodeRed, _
                        True
                        
            ErrorMessage = "Unable to set AF Relays to the correct position for ADwin AF " & ramp_status.Coil & " Ramp. " & _
                           vbNewLine & vbNewLine & _
                           "Would you like to change the ADWIN AF comm settings?" & vbNewLine & _
                           "Cancel = abort AF ramp" & _
                           vbNewLine & vbNewLine & _
                           "Code execution has been paused! " & _
                           "If this AF Ramp is part of a Rockmag / AF run, you will need to Resume the code to " & _
                           "continue the Rockmag / AF run."
            
            UserResp = MsgBox(ErrorMessage, _
                              vbOKCancel, _
                              "ADWIN Relay Position Error")
                              
            frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
            
            Select Case UserResp
            
                Case vbOK
                
                    'Load & show frmSettings and frmADWIN_AF_CommSettings
                    Load frmSettings
                    frmSettings.ZOrder
                    frmSettings.selectTab 4
                    frmSettings.Show
                        
                    'Pop-up message box to remind user to set flow back to running
                    'when done changing the settings
                    MsgBox "Remember to set the program flow back to ""Running"" " & _
                           "in the main window 'Flow' menu or in the Settings form " & _
                           "after you are done editing the ADWIN AF comm settings."
                
                    'Restart the ADWIN ramp
                    DoRampADWIN_WithParameterLogging = DoRampADWIN_WithParameterLogging(MonitorWave, _
                                              UpWave, _
                                              DownWave, _
                                              AF_Data(), _
                                              PeakField, _
                                              HangeTime, _
                                              RampMode, _
                                              RampDownMode, _
                                              DoDCFieldRecord, _
                                              0)
                                              
                    Exit Function
                
                Case Else
                
                    'Exit this sub
                    DoRampADWIN_WithParameterLogging = -1
                    Exit Function
                
            End Select
            
        End If
        
    End If
    
    ramp_status.Ramp_Start_Time = Now
    
    'Start the Ramp process
    ADWIN.Start_Process 1
    
    'Store the start time of the ramp
    StartTime = timeGetTime()
    
    'Pause 100 ms
    PauseTill timeGetTime() + 100
    
    'Loop while checking the ramp process on the ADWIN board every 200 milliseconds
    Do
        
        'Pause 200 ms
        PauseTill timeGetTime() + 200
    
        TempL = ADWIN.Get_Par(4)
    
    Loop Until TempL = 7
    
    ramp_status.Ramp_End_Time = Now
        
    'Get the elapsed time of the ramp
    RampDuration = timeGetTime() - StartTime
    
    PauseTill_NoEvents timeGetTime() + pause_constants.MsecsBetweenRampEndAndReadRampOutputs
    
    'Stop the DC Field record if the user has set for it to be recorded
    'and the record was started successfully
    If DCFieldStatus = True And _
       DoDCFieldRecord = True _
    Then
    
        'Stop the DC Field record
        frm908AGaussmeter.StopDCFieldRecord DCFieldWave, True
        
    End If
           
    'Pause an additional 200 ms
    PauseTill timeGetTime() + 200
            
    'Now Get the basic parameters that we need to know about the AF Ramp
    'whether or not we're in debug or clipping mode
    
   'Get the final point of the DAC output (OUTCOUNT)
    DownWave.CurrentPoint = ADWIN.Get_Par(5)
    
    'Clip off 10 points from the end of the data set
    DownWave.CurrentPoint = DownWave.CurrentPoint - 10
    
    ramp_outputs.Total_Output_Points.TrySetValue (DownWave.CurrentPoint)
    
    'Get the point # of the last ADC input point from the INCOUNT parameter
    MonitorWave.CurrentPoint = ADWIN.Get_Par(6)
    
    'Clip off 10 points from the end of the data set
    MonitorWave.CurrentPoint = MonitorWave.CurrentPoint - 10
    ramp_outputs.Total_Monitor_Points.TrySetValue (MonitorWave.CurrentPoint)

    'Get the point # of the last DAC output point from the Ramp Up process
    UpWave.CurrentPoint = ADWIN.Get_Par(7)
    ramp_outputs.Ramp_Up_Last_Point.TrySetValue (UpWave.CurrentPoint)
    
    'Get the first point # of the DAC output Ramp Down process
    DownWave.StartPoint = ADWIN.Get_Par(8)
    ramp_outputs.Ramp_Down_First_Point.TrySetValue (DownWave.StartPoint)
    
    'Get the max input voltage reached
    MonitorWave.CurrentVoltage = ADWIN.Get_Fpar(4)
    ramp_outputs.Measured_Peak_Monitor_Voltage.TrySetValue (MonitorWave.CurrentVoltage)
    
    'Get the max output voltage reached
    UpWave.CurrentVoltage = ADWIN.Get_Fpar(5)
    ramp_outputs.Max_Ramp_Voltage_Used.TrySetValue (UpWave.CurrentVoltage)
    
    'Get the Down slope (may now be different depending on the ramp mode)
    DownWave.Slope = ADWIN.Get_Fpar(32)
    ramp_outputs.Actual_Slope_Down_Used.TrySetValue (DownWave.Slope)
    
    'Get the Ramp process Time steps from the ADWIN board - ACOUT_TIMESTEP
    UpWave.TimeStep = ADWIN.Get_Fpar(6)
    MonitorWave.TimeStep = UpWave.TimeStep
    ramp_outputs.Time_Step_Between_Points.TrySetValue (UpWave.TimeStep)
    
    'Get the actual used points per period from the ADWIN board - NSAMPLES
    MonitorWave.PtsPerPeriod = ADWIN.Get_Fpar(7)
    ramp_outputs.Number_Points_Per_Period.TrySetValue (MonitorWave.PtsPerPeriod)
    
    If ActiveCoilSystem = AxialCoilSystem Then
    
        ramp_outputs.Coil = "Axial"
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        ramp_outputs.Coil = "Transverse"
        
    Else
    
        ramp_outputs.Coil = "unknown"
        
    End If
    
    Dim zero_threshold As Double
    Dim noise_threshold As Double
    
    
    zero_threshold = 4 * NoiseLevel * _
                     (MonitorWave.range.MaxValue - MonitorWave.range.MinValue) / 2 ^ 16
    
    noise_threshold = 0.005 * ramp_inputs.Peak_Monitor_Voltage.ParamSingle
    
    If noise_threshold < zero_threshold Then
    
        noise_threshold = zero_threshold
        
    End If
    
    
    ramp_status.WasSuccessful = True
    
    If ramp_outputs.Measured_Peak_Monitor_Voltage.ParamSingle < zero_threshold Then
    
        'Zero MOnitor Voltage error
        ramp_status.WasSuccessful = False
        ramp_status.ErrorType = ZeroMonitorVoltage
        
        adwin_af_error.CoilType = CoilType
        adwin_af_error.ErrorType = ZeroMonitorVoltage
        
        adwin_af_error.Source = "frmADWIN_AF.DoRampADWIN_WithParameterLogging"
        adwin_af_error.TimeStamp = Now
        
        modAF_DAQ.UpdateAFErrorStatus adwin_af_error
        
    ElseIf (((ramp_inputs.Peak_Monitor_Voltage.ParamSingle - _
              ramp_outputs.Measured_Peak_Monitor_Voltage.ParamSingle) > noise_threshold) And _
                                       ramp_inputs.ramp_mode.ParamLong < 3) Then
                                       
        'Undershoot MOnitor Voltage error - only for MONITORED ramps
        ramp_status.WasSuccessful = False
        ramp_status.ErrorType = TargetUndershoot
        
        adwin_af_error.CoilType = CoilType
        adwin_af_error.ErrorType = TargetUndershoot
        
        adwin_af_error.Source = "frmADWIN_AF.DoRampADWIN_WithParameterLogging"
        adwin_af_error.TimeStamp = Now
        
        modAF_DAQ.UpdateAFErrorStatus adwin_af_error
                                       
    ElseIf (((ramp_outputs.Measured_Peak_Monitor_Voltage.ParamSingle - _
              ramp_inputs.Peak_Monitor_Voltage.ParamSingle) > noise_threshold) And _
                                       ramp_inputs.ramp_mode.ParamLong < 3) Then
                                       
        'Overshoot monitor voltage error - only for MONITORED ramps
        ramp_status.WasSuccessful = False
        ramp_status.ErrorType = TargetOvershoot
        
        adwin_af_error.CoilType = CoilType
        adwin_af_error.ErrorType = TargetOvershoot
        
        adwin_af_error.Source = "frmADWIN_AF.DoRampADWIN_WithParameterLogging"
        adwin_af_error.TimeStamp = Now
        
        modAF_DAQ.UpdateAFErrorStatus adwin_af_error
                                       
    End If
                                       
    modLogAFParameters.LogAFRamp ramp_inputs, ramp_outputs, ramp_status
    
    'Save data to file if requested
    If RampMode > 1 Then
    
        On Error GoTo Error_Saving_Data
        
        Dim SineFit_Data() As Double
        
        FetchAFData_FromAdwin AF_Data(), _
                              ramp_outputs
                                                    
        If Me.chkVerbose.value = Checked Or _
           modConfig.EnableAFAnalysis Then
           
            SaveAFData AF_Data(), _
                       SineFit_Data(), _
                       ramp_inputs, _
                       ramp_outputs, _
                       ramp_status
                       
        End If
                   
        On Error GoTo 0
        
        GoTo Data_Save_Done
                
Error_Saving_Data:
                   
        frmProgram.StatusBar "Error saving AF Data", 3
        PauseTill_NoEvents timeGetTime() + 2000
                   
    End If
    
Data_Save_Done:
    
    If RampMode = 3 Then
    
        'Save max monitor voltage to MonitorWave
        'Save max ramp voltage to UpWave
        
        MonitorWave.CurrentVoltage = ramp_outputs.Measured_Peak_Monitor_Voltage.ParamSingle
        UpWave.CurrentVoltage = ramp_outputs.Max_Ramp_Voltage_Used.ParamSingle
    
    End If
    
    'Quick error check
    'If the monitor max voltage is < 4 * Noiselevel * 20 / 2^16, then something's wrong with
    '(i.e. if NoiseLevel = 5, then the allowed minimum observed value = 4 * 5 * 3.05e-4 = 6.1 mV
    ' and the Ramp Voltage is at or above the suggested peak, then
    'the monitor channel settings are wrong or somehow messed up.
    With MonitorWave
    
    
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
'
'   FOR DEBUGGING ONLY - Turned off error handling for ADwin AF Ramp
'   All error conditions will now be ignored.
'
'   Original:
'
   If Not ramp_status.WasSuccessful Then
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
'   Changed to:
'
'      If False Then
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
            Dim code_yellow_response As AF_Error_Response
            Dim code_red_response As AF_Error_Response
            
            Set code_yellow_response = _
                modAF_DAQ.AF_Error_Response_Table.GetErrorResponse(adwin_af_error, modStatusCode.CodeYellow)
                
            Set code_red_response = _
                modAF_DAQ.AF_Error_Response_Table.GetErrorResponse(adwin_af_error, modStatusCode.CodeRed)
        
            If RetryNumber < 3 And (code_yellow_response.ErrorAction = ExpressError) Then
            
                PauseForUndershoot ramp_inputs, ramp_outputs, ramp_status, zero_threshold
            
                DoRampADWIN_WithParameterLogging = RetryADWINRamp(MonitorWave, _
                                                                  UpWave, _
                                                                  DownWave, _
                                                                  AF_Data(), _
                                                                  PeakField, _
                                                                  HangeTime, _
                                                                  RampMode, _
                                                                  RampDownMode, _
                                                                  DoDCFieldRecord, _
                                                                  RetryNumber)
                                                                  
                'Exit the function if the retry code returns an error.
                'This skips the portion of the code that reads logged data on the ADwin board
                If DoRampADWIN_WithParameterLogging <> 0 Then
                
                    'Clear all processes on the ADWIN board
                    ReturnVal = ADWIN.ClearAll_Processes
                    
                    PauseTill_NoEvents timeGetTime() + pause_constants.MsecsAfterClearAllProcesses
                   
                    Exit Function
                    
                End If
                                                                 
            ElseIf code_red_response.ErrorAction = ExpressError Then
            
                'No more retries left, set code red error and pause code
                
                Flow_Pause
                SetCodeLevel CodeRed
                
                ErrorMessage = "Monitor input voltage did not reach target monitor voltage on AF " & .BoardUsed.BoardName & _
                               " board.  Tried to re-run AF Ramp three times and encountered same error." & vbNewLine & _
                               "Target Monitor Voltage: " & Format(.PeakVoltage, "#0.000") & vbNewLine & _
                               "Max Monitor Voltage Reached: " & Format(.CurrentVoltage, "#0.000") & _
                               vbNewLine & vbNewLine & _
                               "Code execution has been paused.  Please come and check the machine."
                
                'Something's wrong with the ADWIN board input voltage channel
                'Raise a code red error and send an email
                frmSendMail.MailNotification _
                            "AF Monitor Error after three attempts.", _
                            ErrorMessage, _
                            CodeRed, _
                            True
                            
                ErrorMessage = "Monitor input voltage did not reach target monitor voltage on AF " & .BoardUsed.BoardName & _
                               " board.  Tried to re-run AF Ramp three times and encountered same error.  " & _
                               "Ramp voltage output and/or monitor voltage input comm settings " & _
                               "may be wrong." & vbNewLine & vbNewLine & _
                               "Would you like to change the ADWIN AF comm settings?" & vbNewLine & _
                               "Cancel = abort AF ramp" & vbNewLine & _
                               "No = continue Paleomag code without changing the settings."
                
                UserResp = MsgBox(ErrorMessage, _
                                  vbYesNoCancel, _
                                  "ADWIN Comm Error?")
                                  
                frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
                
                If UserResp = vbYes Then
                
                    RetryNumber = 0
                
                    'Load & show frmSettings and frmADWIN_AF_CommSettings
                    Load frmSettings
                    frmSettings.ZOrder
                    frmSettings.selectTab 4
                    frmSettings.Show
                        
                    'Pop-up message box to remind user to set flow back to running
                    'when done changing the settings
                    MsgBox "Remember to set the program flow back to ""Running"" " & _
                           "in the main window 'Flow' menu or in the Settings form " & _
                           "after you are done editing the ADWIN AF comm. settings."
                
                    'Restart the ADWIN ramp
                    DoRampADWIN_WithParameterLogging = DoRampADWIN_WithParameterLogging(MonitorWave, _
                                              UpWave, _
                                              DownWave, _
                                              AF_Data(), _
                                              PeakField, _
                                              HangeTime, _
                                              RampMode, _
                                              RampDownMode, _
                                              DoDCFieldRecord, _
                                              0)
                
                ElseIf UserResp = vbCancel Then
                
                    'Abort the AF ramp
                    DoRampADWIN_WithParameterLogging = -1
                    
                End If
                
            End If
            
        Else
        
            'No error occurred, return value = 0
            DoRampADWIN_WithParameterLogging = 0
            
        End If
        
    End With
            
    'Clear all processes on the ADWIN board
    ReturnVal = ADWIN.ClearAll_Processes
    
    PauseTill_NoEvents timeGetTime() + pause_constants.MsecsAfterClearAllProcesses
        
    'Reset the Program form status bar to null string
    frmProgram.StatusBar vbNullString, 3
        
    'Deallocate the DC field record wave
    Set DCFieldWave = Nothing
    
End Function



Public Sub DoSineFitAnalysis _
    (ByRef MonitorWave As Wave, _
     ByRef AFMonitor() As Double, _
     ByRef SineFit_Data() As Double, _
     ByVal PeakField As Double, _
     Optional ByVal PtsBetweenFit As Long = 4000)

    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim NumSineFits As Long
    Dim FitLength As Long
    Dim TempL As Long
    
    Dim PercentDone As String
        
    Dim BadFit As Boolean
        
    Dim SineData() As Double
    Dim Sine_Fit() As Double
    Dim Sine_Res() As Double
    Dim FitParams(4) As Double
    Dim RMS As Double
    Dim TempD As Double
    Dim TempD2 As Double
    Dim MaxAmpIn As Double
    Dim MaxAmpOut As Double
    
'    Dim fso As FileSystemObject
'    Dim SineStream As TextStream
    
    'Set Percent Done = 0% padded with two characters
    PercentDone = "  0%"
    
    'Update the Program window status bar with the "Analyzing...__%" info
    frmProgram.StatusBar "Analyzing... " & PercentDone, 3
    
    'Find number of AF monitor input points, and then calculate the number of SineFits
    'to do from that
    NumSineFits = MonitorWave.CurrentPoint \ PtsBetweenFit
    
    'Set N = number of elements in AFMonitor
    N = UBound(AFMonitor, 1)
    
    'Now check to see how many columns AFMonitor has
    TempL = UBound(AFMonitor, 2)
    
    'Redimension the SineFit_Data array
    'If this was a field calibrated ramp, then need an addional column
    If TempL = 3 Then
    
        'This was a field calibrated ramp
        ReDim SineFit_Data(NumSineFits, 11)
        
    Else
    
        'This was an uncalibrated ramp
        ReDim SineFit_Data(NumSineFits, 10)
        
    End If
        
    'Calculate how many points per sine fit from the period of the sine wave + the timestep
    FitLength = (CLng(1 / (MonitorWave.SineFreqMin * MonitorWave.TimeStep)) + 1) * 2
    
    'Re-adjust fit-length if it's greater than the number of points between fits - no overlapping
    'fit windows!
    If FitLength > PtsBetweenFit Then FitLength = PtsBetweenFit
    
    If NumSineFits = 1 And MonitorWave.CurrentPoint < FitLength Then
    
        For j = 0 To 9
        
            SineFit_Data(0, j) = -1
            
        Next j
    
        Exit Sub
        
    End If
    
    'Redimension the Sine Fit function arrays based on FitLength
    ReDim SineData(FitLength)
    ReDim Sine_Fit(FitLength)
    ReDim Sine_est(FitLength)
    
''-------------------------------------------------------------------------------------'
''-------------------------------------------------------------------------------------'
''
''       Debug code only
''
''-------------------------------------------------------------------------------------'
'
'    Set fso = New FileSystemObject
'    Set SineStream = fso.CreateTextFile("C:\Documents and Settings\lab\" & _
'                                        "Desktop\Test MCC Board 11-16-2009\" & _
'                                        "ADWIN Ramp Data\SineDebug_" & _
'                                        Format(Now, "MM-DD-YY_HH-MM-SS") & ".csv")
''-------------------------------------------------------------------------------------'
''-------------------------------------------------------------------------------------'
                                        
    'Loop through the number of sine fits
    For i = 0 To NumSineFits - 1
    
        'Reset MaxAmpIn = 0
        MaxAmpIn = 0
        MaxAmpOut = 0
    
        'Load the correct number of points into the SineData array
        For j = 0 To FitLength - 1
        
            If PtsBetweenFit * i + j > N - 1 Then
            
                'Code is trying to reach beyond the end of the AFMonitor array
                'end for loop
                
                'Reset the number of sine-fits - there were too many specified
                NumSineFits = i
                                
                'Resize the Sine Fit array one element smaller
                ReDim Preserve SineFit_Data(NumSineFits, 10)
            
                'End the sine-fit analysis now - we're out of data points
                Exit Sub
                
            End If
        
            'Set AF Monitor = TempD local variable
            TempD = AFMonitor(PtsBetweenFit * i + j, 0)
            TempD2 = AFMonitor(PtsBetweenFit * i + j, 1)
            
            'Snatch the correct point from the input data array
            SineData(j) = TempD
            
'            Debug.Print SineData(j)
            
            'Find Max IO Voltages during fit period
            If MaxAmpIn < Abs(TempD) Then MaxAmpIn = Abs(TempD)
            If MaxAmpOut < Abs(TempD2) Then MaxAmpOut = Abs(TempD2)
            
        Next j
        
        'Turn on error handling
        On Error Resume Next
        
            'Do the Sine fit
            SineFit SineData, _
                    MonitorWave.TimeStep, _
                    MonitorWave.SineFreqMin, _
                    FitParams, _
                    Sine_Fit, _
                    Sine_Res, _
                    RMS ',
                    'SineStream
                    
            'Check for errors
            If Err.number <> 0 Then
            
                'Sine fit failed, check for inversion error
                'Toggle Fit Error Status Flag
                BadFit = True
                
            Else
            
                BadFit = False
                
            End If
        
        'Return error flow to normal
        On Error GoTo 0
        
        If BadFit = True Then
        
            SineFit_Data(i, 0) = PtsBetweenFit * i
            SineFit_Data(i, 1) = MonitorWave.TimeStep * PtsBetweenFit * i
            SineFit_Data(i, 2) = MaxAmpIn
            SineFit_Data(i, 3) = -1
            SineFit_Data(i, 4) = MonitorWave.SineFreqMin
            SineFit_Data(i, 5) = -1
            SineFit_Data(i, 6) = -1
            SineFit_Data(i, 7) = -1
            SineFit_Data(i, 8) = -1
            SineFit_Data(i, 9) = MaxAmpOut
            
            'Check to see if we also need to store a fake field value
            If TempL = 3 Then
            
                'If AFMonitor has three columns, then SineFit_Data has 11
                SineFit_Data(i, 10) = -1
                
            End If
                
        Else
        
            SineFit_Data(i, 0) = PtsBetweenFit * i
            SineFit_Data(i, 1) = MonitorWave.TimeStep * PtsBetweenFit * i
            SineFit_Data(i, 2) = MaxAmpIn
            SineFit_Data(i, 3) = FitParams(1)
            SineFit_Data(i, 4) = MonitorWave.SineFreqMin
            SineFit_Data(i, 5) = FitParams(2)
            SineFit_Data(i, 6) = FitParams(0)
            SineFit_Data(i, 7) = FitParams(3)
            SineFit_Data(i, 8) = RMS
            SineFit_Data(i, 9) = MaxAmpOut
            
            'Check to see if we also need to calculate & store the field value
            'using the ratio of the Peak field to the peak monitor voltage
            If TempL = 3 Then
            
                'If AFMonitor has three columns, then SineFit_Data has 11
                SineFit_Data(i, 10) = FitParams(1) * PeakField / MonitorWave.PeakVoltage
                
            End If
            
        End If
        
        'Update the data analysis status
        PercentDone = Trim(str(CLng(i / NumSineFits * 100)))
        
        'Pad with whitespace characters so that the length of
        'the whitespace + percentage value = 3 characters
        'then attach a % character
        PercentDone = PadLeft(PercentDone, 4) & "%"
        
'        Debug.Print Trim(Str(i)) & "..." & PercentDone
        
        'Update the status bar
        frmProgram.StatusBar "Analyzing..." & PercentDone, 3
                
        DoEvents
        
    Next i

End Sub

Public Sub DoSineFitAnalysis_UsingAdwinRampClassInstances _
    (ByRef ramp_inputs As AdwinAfInputParameters, _
     ByRef ramp_outputs As AdwinAfOutputParameters, _
     ByRef AFMonitor() As Double, _
     ByRef SineFit_Data() As Double, _
     Optional ByVal PtsBetweenFit As Long = 4000)

    Dim i As Long
    Dim j As Long
    Dim N As Long
    Dim NumSineFits As Long
    Dim FitLength As Long
    Dim TempL As Long
    
    Dim PercentDone As String
        
    Dim BadFit As Boolean
        
    Dim SineData() As Double
    Dim Sine_Fit() As Double
    Dim Sine_Res() As Double
    Dim FitParams(4) As Double
    Dim RMS As Double
    Dim TempD As Double
    Dim TempD2 As Double
    Dim MaxAmpIn As Double
    Dim MaxAmpOut As Double
    
'    Dim fso As FileSystemObject
'    Dim SineStream As TextStream
    
    'Set Percent Done = 0% padded with two characters
    PercentDone = "  0%"
    
    'Update the Program window status bar with the "Analyzing...__%" info
    frmProgram.StatusBar "Analyzing... " & PercentDone, 3
    
    'Find number of AF monitor input points, and then calculate the number of SineFits
    'to do from that
    NumSineFits = ramp_outputs.Total_Monitor_Points.ParamLong \ PtsBetweenFit
    
    'Set N = number of elements in AFMonitor
    N = UBound(AFMonitor, 1)
    
    'This was an uncalibrated ramp
    ReDim SineFit_Data(NumSineFits, 10)
        
    'Calculate how many points per sine fit from the period of the sine wave + the timestep
    FitLength = (CLng(1 / (ramp_inputs.Resonance_Freq.ParamSingle * _
                           ramp_outputs.Time_Step_Between_Points.ParamSingle)) + 1) * 2
    
    'Re-adjust fit-length if it's greater than the number of points between fits - no overlapping
    'fit windows!
    If FitLength > PtsBetweenFit Then FitLength = PtsBetweenFit
    
    If NumSineFits = 1 And ramp_outputs.Total_Monitor_Points.ParamLong < FitLength Then
    
        For j = 0 To 9
        
            SineFit_Data(0, j) = -1
            
        Next j
    
        Exit Sub
        
    End If
    
    'Redimension the Sine Fit function arrays based on FitLength
    ReDim SineData(FitLength)
    ReDim Sine_Fit(FitLength)
    ReDim Sine_est(FitLength)
    
''-------------------------------------------------------------------------------------'
''-------------------------------------------------------------------------------------'
''
''       Debug code only
''
''-------------------------------------------------------------------------------------'
'
'    Set fso = New FileSystemObject
'    Set SineStream = fso.CreateTextFile("C:\Documents and Settings\lab\" & _
'                                        "Desktop\Test MCC Board 11-16-2009\" & _
'                                        "ADWIN Ramp Data\SineDebug_" & _
'                                        Format(Now, "MM-DD-YY_HH-MM-SS") & ".csv")
''-------------------------------------------------------------------------------------'
''-------------------------------------------------------------------------------------'
                                        
    'Loop through the number of sine fits
    For i = 0 To NumSineFits - 1
    
        'Reset MaxAmpIn = 0
        MaxAmpIn = 0
        MaxAmpOut = 0
    
        'Load the correct number of points into the SineData array
        For j = 0 To FitLength - 1
        
            If PtsBetweenFit * i + j > N - 1 Then
            
                'Code is trying to reach beyond the end of the AFMonitor array
                'end for loop
                
                'Reset the number of sine-fits - there were too many specified
                NumSineFits = i
                                
                'Resize the Sine Fit array one element smaller
                ReDim Preserve SineFit_Data(NumSineFits, 10)
            
                'End the sine-fit analysis now - we're out of data points
                Exit Sub
                
            End If
        
            'Set AF Monitor = TempD local variable
            TempD = AFMonitor(PtsBetweenFit * i + j, 0)
            TempD2 = AFMonitor(PtsBetweenFit * i + j, 1)
            
            'Snatch the correct point from the input data array
            SineData(j) = TempD
            
'            Debug.Print SineData(j)
            
            'Find Max IO Voltages during fit period
            If MaxAmpIn < Abs(TempD) Then MaxAmpIn = Abs(TempD)
            If MaxAmpOut < Abs(TempD2) Then MaxAmpOut = Abs(TempD2)
            
        Next j
        
        'Turn on error handling
        On Error Resume Next
        
            'Do the Sine fit
            SineFit SineData, _
                    ramp_outputs.Time_Step_Between_Points.ParamSingle, _
                    ramp_inputs.Resonance_Freq.ParamSingle, _
                    FitParams, _
                    Sine_Fit, _
                    Sine_Res, _
                    RMS ',
                    'SineStream
                    
            'Check for errors
            If Err.number <> 0 Then
            
                'Sine fit failed, check for inversion error
                'Toggle Fit Error Status Flag
                BadFit = True
                
            Else
            
                BadFit = False
                
            End If
        
        'Return error flow to normal
        On Error GoTo 0
        
        If BadFit = True Then
        
            SineFit_Data(i, 0) = PtsBetweenFit * i
            SineFit_Data(i, 1) = ramp_outputs.Time_Step_Between_Points.ParamSingle * PtsBetweenFit * i
            SineFit_Data(i, 2) = MaxAmpIn
            SineFit_Data(i, 3) = -1
            SineFit_Data(i, 4) = ramp_inputs.Resonance_Freq.ParamSingle
            SineFit_Data(i, 5) = -1
            SineFit_Data(i, 6) = -1
            SineFit_Data(i, 7) = -1
            SineFit_Data(i, 8) = -1
            SineFit_Data(i, 9) = MaxAmpOut
                
        Else
        
            SineFit_Data(i, 0) = PtsBetweenFit * i
            SineFit_Data(i, 1) = ramp_outputs.Time_Step_Between_Points.ParamSingle * PtsBetweenFit * i
            SineFit_Data(i, 2) = MaxAmpIn
            SineFit_Data(i, 3) = FitParams(1)
            SineFit_Data(i, 4) = ramp_inputs.Resonance_Freq.ParamSingle
            SineFit_Data(i, 5) = FitParams(2)
            SineFit_Data(i, 6) = FitParams(0)
            SineFit_Data(i, 7) = FitParams(3)
            SineFit_Data(i, 8) = RMS
            SineFit_Data(i, 9) = MaxAmpOut
            
        End If
        
        'Update the data analysis status
        PercentDone = Trim(str(CLng(i / NumSineFits * 100)))
        
        'Pad with whitespace characters so that the length of
        'the whitespace + percentage value = 3 characters
        'then attach a % character
        PercentDone = PadLeft(PercentDone, 4) & "%"
        
'        Debug.Print Trim(Str(i)) & "..." & PercentDone
        
        'Update the status bar
        frmProgram.StatusBar "Analyzing..." & PercentDone, 3
                
        DoEvents
        
    Next i

End Sub


Public Sub ExecuteRamp(ByVal AFCoilSystem As Long, _
                       ByVal PeakValue As Double, _
                       Optional ByVal UpSlope As Double = -1, _
                       Optional ByVal DownSlope As Double = -1, _
                       Optional ByVal IORate As Long = -1, _
                       Optional ByVal PeakHangTime As Double = -1, _
                       Optional ByVal CalRamp As Boolean = True, _
                       Optional ByVal ClipTest As Boolean = False, _
                       Optional ByVal Verbose As Boolean = False, _
                       Optional ByVal DoDCFieldRecord As Boolean = False)
                       
    Dim Freq As Double
    Dim BiggerFreq As Double
    Dim SmallerFreq As Double
    Dim AF_Data() As Double
    Dim SineFit_Data() As Double
    
    Dim FolderName As String
    Dim CoilString As String
    Dim CurTime
    
    Dim ErrorMessage As String
    Dim ErrorCode As Long
    Dim TempL As Long
    Dim RampDownMode As Long
    
   On Error GoTo ExecuteRamp_Error

    If ramp_in_progress Then Exit Sub
    
    ramp_in_progress = False
    
    'Exit if NOCOMM_MODE is on
    If NOCOMM_MODE Then Exit Sub
    
    If EnableAF = False Then
    
        'Msgbox the user, tell them the AF module is disabled
        MsgBox "AF Module is currently disabled.  AF Ramp will now abort.", , _
               "Whoops!"
               
        Exit Sub
        
    End If
    
    ramp_in_progress = True
    
    'Update the program Status bar
    frmProgram.StatusBar "AF Config", 3

    'Based on AFCoilSystem, set the active coil system
    'and set optCoil radio button on frmADWIN_AF
    'The optCoil Click routine will set the frequency to use
    ActiveCoilSystem = AFCoilSystem
    Select Case ActiveCoilSystem
    
        Case AxialCoilSystem
        
            frmADWIN_AF.optCoil(0).value = True
            If Not frmADWIN_AF.TrySetRelays_ADwin(coil_type.Axial) Then
            
                ramp_in_progress = False
                CoilsLocked = False
                Me.chkLockCoils.value = Unchecked
                Exit Sub
                
            End If
        
            CoilString = "AF Axial"
            
        Case TransverseCoilSystem
        
            frmADWIN_AF.optCoil(1).value = True
            If Not frmADWIN_AF.TrySetRelays_ADwin(coil_type.Transverse) Then
            
                ramp_in_progress = False
                CoilsLocked = False
                Me.chkLockCoils.value = Unchecked
                Exit Sub
                
            End If
                        
            CoilString = "AF Transverse"
                    
        Case Else
        
            'No AF coil selected, exit this sub!
            ramp_in_progress = False
            CoilsLocked = False
            Me.chkLockCoils.value = Unchecked
            Exit Sub
        
    End Select
    
    'Lock the coils
    CoilsLocked = True
    Me.chkLockCoils.value = Checked
    
    'Check to see if the AF wave-forms were successfully loaded from the .INI settings file
    If WaveForms Is Nothing Or WaveForms.Count = 0 Then
    
        Flow_Pause
        SetCodeLevel CodeRed
        
        Dim af_error As af_ramp_error
        Set af_error = New af_ramp_error
        
        af_error.CoilType = GetCoilType_FromActiveAFAystem(AFCoilSystem)
        af_error.ErrorType = FatalError
                
        ErrorMessage = "Bad AF Ramp Settings!" & vbNewLine & _
                       "Garbage values must have been dumped into the System WaveForms collection." & _
                       "Please check the Paleomag.ini file for formatting errors." & _
                       vbNewLine & "Code execution has been paused.  Please Come check the machine."
    
        af_error.Message = ErrorMessage
        af_error.Source = "frmADWIN_AF.ExecuteRamp"
        af_error.TimeStamp = Now
        
        modAF_DAQ.UpdateAFErrorStatus af_error
    
        'CRAP!
        frmSendMail.MailNotification _
                    "Program Settings Error!", _
                    ErrorMessage, _
                    CodeRed, _
                    True
                    
        ErrorMessage = "Garbage values must have been dumped into the System WaveForms collection." & _
                       "Please check the Paleomag.ini file for formatting errors." & _
                       vbNewLine & "Code execution will terminate after you click 'OK'"
                       
        MsgBox ErrorMessage, vbCritical, "Fatal AF Error!"

        'This code line is totally pointless given the context, but it's
        'fun to write in anyways. =:-)
        frmProgram.SetProgramCodeLevel StatusCodeColorLevelPrior
        
        End
           
    End If
    
    
    'Check to see if this is a calibrated ramp
    If CalRamp = True Then
    
        'Set the clip-text (unmonitored ramp) check-box to unchecked
        frmADWIN_AF.chkClippingTest.value = Unchecked
    
        'Inputed value will be treated as though it is a Peak Field value
        frmADWIN_AF.txtPeakField = Format(PeakValue, "#0.0##")
        
        'Select the calibrated ramp radio button
        'and deselect the uncalibrated ramp radio button
        frmADWIN_AF.optCalRamp(0).value = True
        frmADWIN_AF.optCalRamp(1).value = False
        
        'Set Ramp Down mode = 1 (use # periods instead of voltage / sec slope)
        RampDownMode = 1
        
    ElseIf CalRamp = False And _
           ClipTest = False _
    Then
           
        'Set the clip-text (unmonitored ramp) check-box to unchecked
        frmADWIN_AF.chkClippingTest.value = Unchecked
           
        'This is an uncalibrated, monitored ramp,
        'Peak value is a Monitor Peak Voltage
        frmADWIN_AF.txtMonitorTrigVolt = Format(PeakValue, "#0.0#####")
        
        'Select the uncalibrated ramp radio button
        'and deselect the calibrated ramp radio button
        frmADWIN_AF.optCalRamp(1).value = True
        frmADWIN_AF.optCalRamp(0).value = False
        
        'Set Ramp Down mode = 1 (use # periods instead of voltage / sec slope)
        RampDownMode = 1
        
    ElseIf CalRamp = False And _
           ClipTest = True _
    Then
    
        'This is an uncalibrated, unmonitored clip-test ramp
        'Peak value is a Ramp Peak Voltage
        frmADWIN_AF.txtRampPeakVoltage = Format(PeakValue, "#0.0#####")
        
        'Set the clip-text (unmonitored ramp) check-box to checked
        frmADWIN_AF.chkClippingTest.value = Checked
        
        'Set Ramp Down mode = 0 (use slope instead of # periods)
        RampDownMode = 0
        
    End If
    
    'Set the peak values now
    SetPeakValues
    
    'Save the peak values to wave objects
    WaveForms("AFRAMPUP").PeakVoltage = val(Me.txtRampPeakVoltage)
    WaveForms("AFRAMPDOWN").PeakVoltage = val(Me.txtRampPeakVoltage)
    WaveForms("AFMONITOR").PeakVoltage = val(Me.txtMonitorTrigVolt)
    
    'Figure out which of the two coil resonance freq's are larger
    'New to use to rescale the slope up & down for the lower freq coil
    'to make sure that the same number of periods are used for the ramp
    'up and ramp down
    If modConfig.AfAxialResFreq > modConfig.AfTransResFreq Then
    
        BiggerFreq = modConfig.AfAxialResFreq
        SmallerFreq = modConfig.AfTransResFreq
        
    Else
    
        BiggerFreq = modConfig.AfTransResFreq
        SmallerFreq = modConfig.AfAxialResFreq
        
    End If
    
    'Set the frequency to use now
    If ActiveCoilSystem = AxialCoilSystem Then
    
        Freq = modConfig.AfAxialResFreq
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        Freq = modConfig.AfTransResFreq
        
    End If
    
    'Update the wave objects
    WaveForms("AFMONITOR").SineFreqMin = Freq
    WaveForms("AFRAMPUP").SineFreqMin = Freq
    WaveForms("AFRAMPDOWN").SineFreqMin = Freq
    
    'Update the ADWIN AF form
    Me.txtFreq.text = Format(Freq, "#0.0##")
    
    'Determine what the Up & Down slopes should be
    If UpSlope = -1 Then
    
        'No ramp-up slope inputed.
        'Set ramp-up slope depending on the Ramp Peak voltage
        UpSlope = GetUpSlope(WaveForms("AFRAMPUP").PeakVoltage)
        
    End If
        
'    UpSlope = RoundSlopeToPeriod(UpSlope, _
'                                 WaveForms("AFRAMPUp").PeakVoltage, _
'                                 WaveForms("AFRAMPUp").SineFreqMin)
        
    'Set the Slope Up
    WaveForms("AFRAMPUP").Slope = UpSlope
    
    'Update the ADWIN AF form
    Me.txtRampUpSlope = Format(WaveForms("AFRAMPUP").Slope, "#0.0##")
        
    If DownSlope = -1 Then
    
        'No ramp-down slope inputed.
        'Set ramp-down slope depending on the Ramp Peak voltage
        DownSlope = GetDownSlope(WaveForms("AFRAMPDOWN").PeakVoltage)
        
    End If
        
    'Round down slope to the nearest period value
    DownSlope = RoundSlopeToPeriod(DownSlope, _
                                   WaveForms("AFRAMPDOWN").PeakVoltage, _
                                   WaveForms("AFRAMPDOWN").SineFreqMin)
        
    'Set the Slope Down
    WaveForms("AFRAMPDOWN").Slope = DownSlope
        
    'Update the ADWIN AF form
    Me.txtRampDownSlope = Format(WaveForms("AFRAMPDOWN").Slope, "#0.0##")
    
    'Set the IO Rate
    If IORate = -1 Then
    
        'No IORate given, use the default wave object's IORates
        IORate = WaveForms("AFRAMPUP").IORate
            
        'Propagate this IOrate to the two other ADWIN AF wave objects
        WaveForms("AFRAMPDOWN").IORate = IORate
        WaveForms("AFMONITOR").IORate = IORate
                
        'Update the form with this IORate
        Me.txtRampRate = Trim(str(IORate))
        
    Else
        
        'Update all the ADWIN AF wave objects' IORates
        WaveForms("AFRAMPUP").IORate = IORate
        WaveForms("AFRAMPDOWN").IORate = IORate
        WaveForms("AFMONITOR").IORate = IORate
        
        'Update the form
        Me.txtRampRate = Trim(str(IORate))
        
    End If
        
    'Update the TimeSteps of all three ADWIN AF wave objects
    WaveForms("AFRAMPUP").TimeStep = 1 / IORate
    WaveForms("AFRAMPDOWN").TimeStep = 1 / IORate
    WaveForms("AFMONITOR").TimeStep = 1 / IORate
            
    'Check to see if a PeakHangTime was inputed
    If PeakHangTime = -1 Then
    
        'Calculate 100 periods worth of peak hang time in miliseconds
        PeakHangTime = modConfig.HoldAtPeakField_NumPeriods * 1 / Freq * 1000
        
        'Need to prevent to small of a peak hang time (time needs to be > 100 ms)
        If PeakHangTime < 100 Then PeakHangTime = 100
        
        'Update the form
        Me.txtRampPeakDuration = Format(PeakHangTime, "#0.0####")
            
    Else
    
        'Need to prevent to small of a peak hang time
        If PeakHangTime < 100 Then PeakHangTime = 100
    
        'Update the form
        Me.txtRampPeakDuration = Format(PeakHangTime, "#0.0####")
        
    End If
    
    'Set the predicted durations for the various pieces of the ramp cycle
    With WaveForms("AFRAMPUP")
    
        If .PeakVoltage = 0 Then
        
            ramp_in_progress = False
            Exit Sub
            
        End If
        
        .Duration = .PeakVoltage / .Slope * 1000
        
        'Display this duration on the form
        Me.lblRampUpDuration = Trim(str(.Duration))
        
    End With
                       
    With WaveForms("AFRAMPDOWN")
    
        .Duration = .PeakVoltage / .Slope * 1000
        
        'Display this duration on the form
        Me.lblRampDownDuration = Trim(str(.Duration))
        
    End With
    
    'Set the AF Ramp cycle total duration + 200 ms for kicks & giggles
    With WaveForms("AFMONITOR")
    
        .Duration = WaveForms("AFRAMPUP").Duration + _
                    WaveForms("AFRAMPDOWN").Duration + _
                    PeakHangTime + _
                    200
                                      
        'Display this duration on the form
        Me.lblTotalRampDuration = Trim(str(.Duration))
        
    End With
                       
    'Determine the Ramp mode
    If ClipTest = True Then
    
        RampMode = 3
        
    ElseIf Verbose = True Or _
           modConfig.EnableAFAnalysis = True _
    Then
    
        RampMode = 2
        
    Else
    
        RampMode = 1
        
    End If
                       
    'Now do the ramp
    ErrorCode = DoRampADWIN(WaveForms("AFMONITOR"), _
                            WaveForms("AFRAMPUP"), _
                            WaveForms("AFRAMPDOWN"), _
                            AF_Data, _
                            val(Me.txtPeakField), _
                            PeakHangTime, _
                            RampMode, _
                            RampDownMode, _
                            DoDCFieldRecord)
                           
    'Blank the second panel of the status bar
    frmProgram.StatusBar "", 2
                              
    If Not TrySetAllRelaysToDefaultPosition_ADwin Then GoTo SetAllRelays_Down_Error
                              
    'Unlock the coils
    CoilsLocked = False
    Me.chkLockCoils.value = Unchecked
    
    ramp_in_progress = False

   On Error GoTo 0
   Exit Sub

SetAllRelays_Down_Error:

    Flow_Pause
    SetCodeLevel CodeYellow

    ramp_in_progress = False

    ErrorMessage = "AF/IRM Relay Error" & vbNewLine & _
                   "A AF/IRM relay control error occurred after execution of the current AF Ramp.  " & _
                   "Please check the status of the current AF ramp in the Ramp Status Logs." & _
                   vbNewLine & "Code execution will continue."
                  
    frmSendMail.MailNotification _
                "AF/IRM Relay Error", _
                ErrorMessage, _
                CodeYellow
                   
    Flow_Resume
    Exit Sub

ExecuteRamp_Error:
    
    Flow_Pause
    SetCodeLevel CodeRed

    ramp_in_progress = False
    
    Dim af_fatal_error As af_ramp_error
    Set af_fatal_error = New af_ramp_error
    
    af_fatal_error.CoilType = GetCoilType_FromActiveAFAystem(AFCoilSystem)
    af_fatal_error.ErrorType = FatalError
    af_fatal_error.ErrNumber = Err.number
    af_fatal_error.Source = "frmADWIN_AF.ExecuteRamp"
    af_fatal_error.TimeStamp = Now
    
    ErrorMessage = "Unexpected AF Ramp Error" & vbNewLine & _
                   "An unexpected error occurred while running trying to execute the current AF Ramp.  The current AF ramp has been aborted." & _
                   vbNewLine & "Code execution has been paused.  Please Come check the machine." & _
                   vbNewLine & vbNewLine & _
                   "--- Error Details --- " & vbNewLine & _
                   "Number: " & Trim(CStr(Err.number)) & vbNewLine & _
                   "Source: " & Err.Source & vbNewLine & _
                   "Description: " & Err.Description & vbNewLine
                  
    af_fatal_error.Message = ErrorMessage
    
    modAF_DAQ.UpdateAFErrorStatus af_fatal_error
                  
    frmSendMail.MailNotification _
                "AF Ramp Error!", _
                ErrorMessage, _
                CodeRed, _
                True
                   
    MsgBox ErrorMessage, vbCritical, "Fatal AF Error!"
                      
End Sub

Public Sub FetchAFData_FromAdwin(ByRef AF_Data() As Double, _
                                 ByRef ramp_outputs As AdwinAfOutputParameters)

    Dim TempDataIn() As Long
    Dim TempDataOut() As Long

    'Update Status bar again
    frmProgram.StatusBar "Getting data...", 3

    'Pause for 1/4 the duration of the last ramp cycle for Ramp Data arrays
    'to become available
    PauseTill timeGetTime() + (ramp_outputs.GetTotalRampDuration() * 1000) \ 4
    
    'Set N = the maximum # of data points that can be stored by
    'the ADWIN code = MAXALLOWEDDATAPTS
    N = ADWIN.Get_Par(11)

    'Redimension the Temp & AF_Data arrays so that they are the
    'Same size as the number of INCOUNT points
    
    Dim num_points As Long
    num_points = ramp_outputs.Total_Monitor_Points.ParamLong
        
    'If this is a calibrated ramp, need three columns in AF_Data
    ReDim AF_Data(num_points, 2)
    
    ReDim TempDataIn(num_points + 1)
    ReDim TempDataOut(num_points + 1)
    
    'No get the data from the LONG valued ADWIN external memory data arrays
    'loaded into TempDataIn and TempDataOut
    ADWIN.GetData_Long 31, 1, num_points, TempDataIn
    ADWIN.GetData_Long 32, 1, num_points, TempDataOut
    
    'Set Percent done to 0%
    PercentDone = "  0%"
    
    'Data has been retrieved, change status bar status to "Converting..."
    frmProgram.StatusBar "Converting... " & PercentDone, 3
    
    Dim adwin_range As range
    Set adwin_range = New range
    adwin_range.MaxValue = 10
    adwin_range.MinValue = -10
    
    For i = 1 To num_points
    
        AF_Data(i - 1, 0) = adwin_range.ADWIN_RangeConverter(, TempDataIn(i))
        AF_Data(i - 1, 1) = adwin_range.ADWIN_RangeConverter(, TempDataOut(i))
        
        'Every 1 hundred points, update data-conversion status
        If i Mod 100 = 0 Then
        
            'Format the percent done string
            PercentDone = Trim(str(CInt(i / num_points * 100)))
            PercentDone = PadLeft(PercentDone, 4) & "%"
            
            'Update the program form status bar
            frmProgram.StatusBar "Converting... " & PercentDone, 3
            
        End If
        
    Next i

End Sub

Public Function GetCoilType_FromActiveAFAystem(ByVal af_coil_system As Long) As coil_type

    Select Case af_coil_system
    
        Case modConfig.AxialCoilSystem
        
            GetCoilType_FromActiveAFAystem = coil_type.Axial
        
        Case modConfig.TransverseCoilSystem
        
            GetCoilType_FromActiveAFAystem = coil_type.Transverse
        
        Case Else
        
            GetCoilType_FromActiveAFAystem = coil_type.Unknown
            
    End Select

End Function

Public Sub SaveAFData(ByRef AF_Data() As Double, _
                      ByRef SineFit_Data() As Double, _
                      ByRef ramp_inputs As AdwinAfInputParameters, _
                      ByRef ramp_outputs As AdwinAfOutputParameters, _
                      ByRef ramp_status As AdwinAfRampStatus)

    If (ramp_status.WasSuccessful) Then
    
        'Do sine fits on the monitor input data
        DoSineFitAnalysis_UsingAdwinRampClassInstances _
                          ramp_inputs, _
                          ramp_outputs, _
                          AF_Data, _
                          SineFit_Data, _
                          500
                          
    Else
    
        'Set Sine-Fit data to empty value
        ReDim SineFit_Data(1, 1)
        SineFit_Data(0, 0) = -1
        
    End If
        
    Dim FolderName As String
    Dim cur_time As Date
    
    cur_time = Now
    
    FolderName = Get_SaveAFData_FolderName(cur_time, ramp_inputs, ramp_status)
    
    'Now Call the Ramp data save function
    frmFileSave.MultiRampFileSave_ADWIN AF_Data(), _
                                        ramp_outputs.Time_Step_Between_Points.ParamSingle, _
                                        1048000, _
                                        FolderName, _
                                        cur_time, _
                                        SineFit_Data(), _
                                        False, _
                                        True, _
                                        CLng(ramp_outputs.Number_Points_Per_Period.ParamSingle \ 2) + 1
   
End Sub

Public Function Get_SaveAFData_FolderName(ByVal current_time As Date, _
                                          ByRef ramp_inputs As AdwinAfInputParameters, _
                                          ByRef ramp_status As AdwinAfRampStatus) As String

    'Set the Local Data folder name
    'Check the Ramp Mode
    
    FolderName = ramp_inputs.GetShortRampDescrip()
    
    If Not ramp_status.WasSuccessful Then
    
        FolderName = FolderName & " (" & ramp_status.GetErrorShortType_AsString & ")"
        
    End If
    
    If Me.optCalRamp(0).value = True Then
    
        'This is a calibrated ramp, label the folder by the voltage being ramped to
        FolderName = FolderName & " " & ramp_status.TargetPeakField & " " & modConfig.AFUnits & " "
    
    ElseIf ramp_inputs.ramp_mode.ParamLong = 3 Then
    
        FolderName = FolderName & " " & Format(ramp_inputs.Peak_Ramp_Voltage.ParamSingle, "#0.0###") & "V "
    
    Else
    
        FolderName = FolderName & " " & Format(ramp_inputs.Peak_Monitor_Voltage.ParamSingle, "#0.0###") & "V "
    
    End If
              
    Get_SaveAFData_FolderName = FolderName & " - " & Format(current_time, "MM-DD-YY, HH MM SS") & "/"

End Function

Private Sub Form_Activate()

    If EnableAF = False Then
        
        'AF's not enabled, cannot Tune the AF coils
        'Tell user that calibration is turned off, but
        'can still edit values
        MsgBox "The AF module is currently disabled.  AF Ramp cycles " & _
               " cannot be performed now." & _
               "Whoops!"
               
        'Disable all the necessary buttons on the form
        Me.cmdStartRamp.Enabled = False
        
    Else
    
        'Enable all the necessary buttons on the form
        Me.cmdStartRamp.Enabled = True
        
    End If

    'First propagate the locked coils state
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked

    'If the window is activated, need to propagate
    'the current active coil settings to the radio buttons
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
    Else
    
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        If AFSystem = "ADWIN" Then
        
            'Set all relays low
            On Error Resume Next
            ADWIN.Set_Digout 63
            On Error GoTo 0
            
        End If
        
    End If

End Sub

Private Sub Form_Load()

   'Set the width to the correct width
    Me.Width = 6975
    Me.Height = 8805
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    Me.Caption = "AF Demag Control"

    If EnableAF = False Then
        
        'AF's not enabled, cannot Tune the AF coils
        
        'Disable all the necessary buttons on the form
        Me.cmdStartRamp.Enabled = False
        
    Else
    
        'Enable all the necessary buttons on the form
        Me.cmdStartRamp.Enabled = True
        
    End If

    'Set isUserChange = True
    isUserChange = True

    'First propagate the locked coils state
    If CoilsLocked = True Then Me.chkLockCoils.value = Checked
    If CoilsLocked = False Then Me.chkLockCoils.value = Unchecked

    'If the window is activated, need to propagate
    'the current active coil settings to the radio buttons
    If ActiveCoilSystem = AxialCoilSystem Then
    
        optCoil(0).value = True
        
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        optCoil(1).value = True
        
        '??? //3/31/17
    ElseIf ActiveCoilSystem = IRMAxialCoilSystem Then
    
        optCoil(2).value = True
    '??? //3/31/17
        
    Else
    
        optCoil(0).value = False
        optCoil(1).value = False
        
        ActiveCoilSystem = NoCoilSystem
        If AFSystem = "ADWIN" Then
        
            'Set all relays low
            On Error Resume Next
            ADWIN.Set_Digout 63
            On Error GoTo 0
            
        End If
        
    End If

    'Set Interval at which timeGetTime() command operates to 1 micro-s
    timeBeginPeriod 1

    'Set Un-monitored /clipping test check-box to unchecked.
    '(Clipping test = unmonitored ramp cycle)
    Me.chkClippingTest.value = Unchecked
   
    'Set Debug (AF Data save) mode to off
    Me.chkVerbose.value = Unchecked
   
'    Debug.Print "1) Active Coil System: " & Trim(Str(ActiveCoilSystem))
   
    'Load Sine Freq Values from Global Settings
    If ActiveCoilSystem = AxialCoilSystem Then
    
        'Display Res Freq
        Me.txtFreq = Trim(str(modConfig.AfAxialResFreq))
                
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        Me.txtFreq = Trim(str(modConfig.AfTransResFreq))
        
    Else
    
        Me.txtFreq = ""
        
    End If
    
    Me.txtRampRate = WaveForms("AFRAMPUP").IORate
    
    'Blank duration labels
    Me.lblRampDownDuration = ""
    Me.lblRampUpDuration = ""
    Me.lblTotalRampDuration = ""
            
'    Debug.Print "2) Active Coil System: " & Trim(Str(ActiveCoilSystem))
            
End Sub

Public Function GetDownSlope(ByVal RampPeakVolts As Double) As Double

    Dim RampDownPeriods As Long
    Dim RampPeriod As Double
    Dim RampDuration As Double

    If ActiveCoilSystem = AxialCoilSystem Then RampPeriod = 1 / modConfig.AfAxialResFreq
    If ActiveCoilSystem = TransverseCoilSystem Then RampPeriod = 1 / modConfig.AfTransResFreq

    'Get the initial calculated ramp-Down duration
    'based Downon the Max Ramp-Down time setting
    RampDownPeriods = CLng(modConfig.RampDownNumPeriodsPerVolt * RampPeakVolts)
    
    If RampDownPeriods < modConfig.MinRampDown_NumPeriods Then
    
        RampDownPeriods = modConfig.MinRampDown_NumPeriods
        
    End If
    
    If RampDownPeriods > modConfig.MaxRampDown_NumPeriods Then
    
        RampDownPeriods = modConfig.MaxRampDown_NumPeriods
        
    End If
    
    'Now calculate the RampDuration from this (ramp duration is in SECONDS)
    RampDuration = RampDownPeriods * RampPeriod
    
    GetDownSlope = RampPeakVolts / RampDuration
    
End Function

'Public Sub GetUpSlope
'
' Created: August 5, 2010
'  Author: Isaac Hilburn
'
' Summary: This function is the brains behind how fast an ADWIN AF ramp up
'          is allowed to happen.
'
'  Inputs:

Public Function GetUpSlope(ByVal RampPeakVolts As Double) As Double

    
    Dim RampUpDur_ms As Long

    'Compare the RampPeakVolts to the Ramp Voltage corresponding to the
    'peak field (if the calibration is done), if not, then relative to the
    'Max ramp voltage set in the AF Auto tune form
    If ActiveCoilSystem = AxialCoilSystem Then
    
        'Need to multiply by 1000 to convert seconds to miliseconds
        RampUpDur_ms = RampPeakVolts / modConfig.AxialRampUpVoltsPerSec * 1000
    
    ElseIf ActiveCoilSystem = TransverseCoilSystem Then
    
        'Need to multiply by 1000 to convert seconds to miliseconds
        RampUpDur_ms = RampPeakVolts / modConfig.TransRampUpVoltsPerSec * 1000
    
    Else
    
        'If no coil is selected, default to a 1 second ramp
        GetUpSlope = RampPeakVolts
        
        Exit Function
        
    End If
    
    'Check to make sure RampUpDur_ms is not smaller than the minimum allowed ramp duration
    If RampUpDur_ms < modConfig.MinRampUpTime_ms Then
    
        RampUpDur_ms = modConfig.MinRampUpTime_ms
        
    End If
    
    If RampUpDur_ms > modConfig.MaxRampUpTime_ms Then
    
        RampUpDur_ms = modConfig.MaxRampUpTime_ms
        
    End If
    
    'Now can use the ramp up duration to calculate the Ramp Up slope
    'Note: RampUpDur_ms is in miliseconds
    GetUpSlope = RampPeakVolts / (RampUpDur_ms / 1000)
        
End Function

Public Function IsCorrect_AFRelay_position(ByVal CoilType As coil_type) As Boolean

    Dim out_byte As Long
    Dim status_byte As Byte
    
    On Error Resume Next
        out_byte = ADWIN.Get_Digout()
        
        If Err.number <> 0 Then GoTo Error_Getting_AF_Relay_Status
    On Error GoTo 0
        
    On Error GoTo Error_Processing_AF_Relay_Status

    If out_byte >= 0 And out_byte < 64 Then
        
        status_byte = out_byte Mod 64
                
    Else
    
        IsCorrect_AFRelay_position = False
        
    End If

    Dim all_up_byte As Byte
    
    all_up_byte = (2 ^ 6) - 1

    Select Case CoilType
    
        Case coil_type.Axial
        
            If status_byte = all_up_byte - (2 ^ (modConfig.AFAxialRelay.ChanNum)) Then
            
                IsCorrect_AFRelay_position = True
                
            Else
                
                IsCorrect_AFRelay_position = False
                
            End If
        
        Case coil_type.Transverse
        
            If status_byte = all_up_byte - (2 ^ (modConfig.AFTransRelay.ChanNum)) Then
            
                IsCorrect_AFRelay_position = True
                
            Else
                
                IsCorrect_AFRelay_position = False
                
            End If
        
        Case coil_type.IRMAxial
        
            If status_byte = (Not (2 ^ (modConfig.AFAxialRelay.ChanNum))) And _
                             (2 ^ (modConfig.AFTransRelay.ChanNum)) Then
            
                IsCorrect_AFRelay_position = True
                
            Else
                
                IsCorrect_AFRelay_position = False
                
            End If
        
        Case coil_type.IRMTrans
        
            If status_byte = (Not (2 ^ (modConfig.AFTransRelay.ChanNum))) And _
                             (2 ^ (modConfig.AFAxialRelay.ChanNum)) Then
            
                IsCorrect_AFRelay_position = True
                
            Else
                
                IsCorrect_AFRelay_position = False
                
            End If
        
        Case Else
        
            IsCorrect_AFRelay_position = True
        
    End Select
    
    On Error GoTo 0
    Exit Function
    
Error_Getting_AF_Relay_Status:

    IsCorrect_AFRelay_position = False
    Exit Function
    
Error_Processing_AF_Relay_Status:

    IsCorrect_AFRelay_position = False
    Exit Function

End Function

Private Sub optCalRamp_Click(Index As Integer)

    'Check to see if this is a user change
    If isUserChange Then
    
        'Use this so the lines of code below don't trigger this event recursively
        isUserChange = False
    
        'If one Calibration option is selected, then deselect the other
        If optCalRamp(0).value = True Then optCalRamp(1).value = False
        If optCalRamp(1).value = True Then optCalRamp(0).value = False
        
    End If
    
    isUserChange = True

End Sub

Public Sub optCoil_Click(Index As Integer)
    
    'Check to see if the coil change / selection is locked
    If CoilsLocked = True Or _
       EnableAF = False _
    Then Exit Sub
    
    If Index = 0 And _
       optCoil(Index).value = True _
    Then
    
        ActiveCoilSystem = AxialCoilSystem
        
        'reset the freq text-box
        Me.txtFreq = Trim(str(modConfig.AfAxialResFreq))
        
        'Activate the frequency change event
        txtFreq_Change
            
    ElseIf Index = 1 And _
       optCoil(Index).value = True _
    Then
    
        ActiveCoilSystem = TransverseCoilSystem
        
        'reset the freq text-box
        Me.txtFreq = Trim(str(modConfig.AfTransResFreq))
        
        'Activate the frequency change event
        txtFreq_Change
        
        '??? //3/31/17
    ElseIf Index = 2 And _
       optCoil(Index).value = True _
    Then
    
        ActiveCoilSystem = IRMAxialCoilSystem
        
        'reset the freq text-box
        Me.txtFreq = Trim(str(modConfig.AfAxialResFreq))
        
        'Activate the frequency change event
        txtFreq_Change
        '??? //3/31/17
    Else
    
        'No coil system set
        ActiveCoilSystem = NoCoilSystem
        
        'reset the freq text-box
        Me.txtFreq = vbullstring
    
    End If
    
End Sub

Public Sub PauseForUndershoot(ByRef ramp_inputs As AdwinAfInputParameters, _
                              ByRef ramp_outputs As AdwinAfOutputParameters, _
                              ByRef ramp_status As AdwinAfRampStatus, _
                              ByVal zero_threshold As Double)
                              
    If ((ramp_inputs.Peak_Monitor_Voltage.ParamSingle _
          - ramp_outputs.Measured_Peak_Monitor_Voltage.ParamSingle) > zero_threshold) _
        And _
        ramp_inputs.ramp_mode.ParamLong < 3 Then
                    
        'An undershoot error has occurred
        'Calculate pause time based on target
        'peak monitor voltage relative to
        'Max allowed peak monitor voltage
        Dim target_ratio_to_max As Double
                
        If ramp_inputs.Max_Monitor_Voltage.ParamSingle <= 0.0001 Then Exit Sub
        
        target_ratio_to_max = ramp_inputs.Peak_Monitor_Voltage.ParamSingle / ramp_inputs.Max_Monitor_Voltage.ParamSingle
        
        Dim maximum_wait_time_seconds As Long
        
        maximum_wait_time_seconds = 300
        
        If target_ratio_to_max > 0.8 Then
        
            modAF_DAQ.PauseBetweenUseCoils_InSeconds maximum_wait_time_seconds * target_ratio_to_max
            
        End If
        
    End If
                              
End Sub

Public Function RetryADWINRamp( _
    ByRef MonitorWave As Wave, _
    ByRef UpWave As Wave, _
    ByRef DownWave As Wave, _
    ByRef AF_Data() As Double, _
    ByVal PeakField As Double, _
    ByVal HangeTime As Long, _
    ByVal RampMode As Long, _
    ByVal RampDownMode As Long, _
    ByVal DoDCFieldRecord As Boolean, _
    ByVal RetryNumber As Integer) As Long

    Dim my_retry_number As Integer
    my_retry_number = RetryNumber + 1
    
    'Update Status bar For Redo
    frmProgram.StatusBar "AF Redo #" & Trim(CStr(my_retry_number)), 2

    'Set code level to yellow
    SetCodeLevel CodeYellow
        
    With MonitorWave
        
        On Error GoTo RetryADWINRamp_SendNotificationEmailError
        
        Dim ErrorMessage As String
        
        ErrorMessage = "Max observed monitor voltage did not reach target monitor voltage on AF " & .BoardUsed.BoardName & _
                       " board. " & vbNewLine & vbNewLine & _
                       "Retrying AF Ramp, attempt #" & Trim(CStr(my_retry_number)) & vbNewLine & vbNewLine & _
                       "Target Monitor Voltage: " & Format(.PeakVoltage, "#0.000") & vbNewLine & _
                       "Max Monitor Voltage Reached: " & Format(.CurrentVoltage, "#0.000") & _
                       vbNewLine & vbNewLine & _
                       "Code execution will continue."
        
        
        
        
        frmSendMail.MailNotification _
                        "Redo AF #" & Trim(CStr(my_retry_number)) & ", after AF Monitor Error", _
                        ErrorMessage, _
                        CodeYellow
                        
        On Error GoTo 0
                        
        
    End With
    
RetryADWINRamp_SendNotificationEmailError:
    
    SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
        
    'Toggle to other coil system
    Select Case modConfig.ActiveCoilSystem
    
        Case AxialCoilSystem
        
            frmADWIN_AF.TrySetRelays_ADwin coil_type.Transverse
            frmADWIN_AF.TrySetRelays_ADwin coil_type.Axial
        
        Case TransverseCoilSystem
        
            frmADWIN_AF.TrySetRelays_ADwin coil_type.Axial
            frmADWIN_AF.TrySetRelays_ADwin coil_type.Transverse
        
        Case Else
        
            'Abort and pass back non-zero value
            RetryADWINRamp = -1
                    
    End Select
        
    'Pause 5 seconds
    modAF_DAQ.PauseBetweenUseCoils_InSeconds 5
    
    RetryADWINRamp = DoRampADWIN_WithParameterLogging( _
                        MonitorWave, _
                        UpWave, _
                        DownWave, _
                        AF_Data, _
                        PeakField, _
                        HangeTime, _
                        RampMode, _
                        RampDownMode, _
                        DoDCFieldRecord, _
                        my_retry_number)
                                                            

    frmProgram.StatusBar "", 2

End Function


Public Function RoundSlopeToPeriod(ByVal Slope As Double, _
                                   ByVal PeakVoltage As Double, _
                                   ByVal SineFreq As Double) As Double
                              
    Dim NumPeriods As Long
    
    'If a zero slope is put in, output a zero slope
    If Slope = 0 Then
    
        RoundSlopeToPeriod = 0
        
        Exit Function
        
    End If
    
    'Slope is non-zero, calculate the number of whole periods closest
    'to that slope
    NumPeriods = CLng(PeakVoltage / Slope * SineFreq)
    
    'Recalculate the slope from NumPeriods
    RoundSlopeToPeriod = PeakVoltage * SineFreq / NumPeriods
                              
End Function

Public Function TrySetAllRelaysDown_ADwin() As Boolean

    Dim TTLBoard As Board
    Dim BoardName As String
    
    Dim relay_byte_id As Byte
        
    Set TTLBoard = Nothing
    
On Error GoTo Wrong_AF_System_Error
    
        'This function is designed only for the ADwin system to control the relays
        If AFSystem <> "ADWIN" Then GoTo Wrong_AF_System_Error
                    
On Error GoTo 0

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then
    
        TrySetAllRelaysDown_ADwin = True
        Exit Function
        
    End If
    
'Turn on Error Handling
On Error GoTo Set_Board_Error
    
    BoardName = AFAxialRelay.BoardName

    'Snag the board associated with the AfAxial Relay Channel
    Set TTLBoard = SystemBoards(AFAxialRelay.BoardName)
                  
   
On Error GoTo 0
        
On Error GoTo Board_Is_Missing
    
    'Error Check again
    If TTLBoard Is Nothing Then GoTo Board_Is_Missing
    
On Error GoTo 0
    
On Error GoTo Boot_Board_Error
    
    'Turn off the error-pop-up in the Boot process
    ADWIN.Show_Errors (0)
    
    'Check to make sure the ADWIN board is booted
    If ADWIN.ADWIN_BootBoard(TTLBoard) = False Then GoTo Boot_Board_Error
    
On Error GoTo 0

On Error GoTo Set_Relay_Error
    
    'Set the relays all down
    ADWIN.Set_Digout 0

    Set TTLBoard = Nothing
    
    'Hard Pause - 1 second - do not allow other events to be handled.
    PauseTill_NoEvents timeGetTime() + 1000
    
On Error GoTo 0
    
    TrySetAllRelaysDown_ADwin = True
    Exit Function
        
Dim error_subject As String
Dim error_message As String
        
Wrong_AF_System_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Down position." & vbNewLine & vbNewLine & _
              "AF System currently not set to ""ADWIN"". " & vbNewLine & _
              "This method is only supported for AF systems that use an ADwin device. " & vbNewLine & _
              "Currently, AF System is set to: " & AFSystem & vbNewLine & vbNewLine & _
              "Please check your AF System settings." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysDown_ADwin = False
    Exit Function
                      
        
Set_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Down position." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysDown_ADwin = False
    Exit Function
    
Board_Is_Missing:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Down position." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysDown_ADwin = False
    Exit Function
    
Boot_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Down position." & vbNewLine & vbNewLine & _
              "Unable to boot the ADWIN device." & _
              vbNewLine & "ADWIN Dev # = " & Trim(str(TTLBoard.BoardNum)) & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysDown_ADwin = False
    Exit Function
    
Set_Relay_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Down position." & vbNewLine & vbNewLine & _
              "One or more calls to set digital output byte on the ADWIN device may have failed." & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysDown_ADwin = False
    Exit Function


End Function

Public Function TrySetAllRelaysUp_ADwin() As Boolean

    Dim TTLBoard As Board
    Dim BoardName As String
        
    Set TTLBoard = Nothing
    
On Error GoTo Wrong_AF_System_Error
    
        'This function is designed only for the ADwin system to control the relays
        If AFSystem <> "ADWIN" Then GoTo Wrong_AF_System_Error
                    
On Error GoTo 0

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then
    
        TrySetAllRelaysUp_ADwin = True
        Exit Function
        
    End If
    
'Turn on Error Handling
On Error GoTo Set_Board_Error
    
    BoardName = AFAxialRelay.BoardName

    'Snag the board associated with the AfAxial Relay Channel
    Set TTLBoard = SystemBoards(AFAxialRelay.BoardName)
                  
   
On Error GoTo 0
        
On Error GoTo Board_Is_Missing
    
    'Error Check again
    If TTLBoard Is Nothing Then GoTo Board_Is_Missing
    
On Error GoTo 0
    
On Error GoTo Boot_Board_Error
    
    'Turn off the error-pop-up in the Boot process
    ADWIN.Show_Errors (0)
    
    'Check to make sure the ADWIN board is booted
    If ADWIN.ADWIN_BootBoard(TTLBoard) = False Then GoTo Boot_Board_Error
    
On Error GoTo 0

On Error GoTo Set_Relay_Error
    
    
    Dim prior_status As String
    prior_status = frmProgram.sbStatusBar.Panels(2).text
    frmProgram.StatusBar "Raising all relays", 2

    'Set all six digital output channels to logic High (a.k.a. relays to the up orientation)
    ADWIN.Set_Digout 63
    
    'Hard Pause - 2 second - do not allow other events to be handled.
    PauseTill_NoEvents timeGetTime() + 2000
    
    Set TTLBoard = Nothing
    
    frmProgram.StatusBar prior_status, 2
    
    
On Error GoTo 0
    
    TrySetAllRelaysUp_ADwin = True
    Exit Function
        
Dim error_subject As String
Dim error_message As String
        
Wrong_AF_System_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Up position." & vbNewLine & vbNewLine & _
              "AF System currently not set to ""ADWIN"". " & vbNewLine & _
              "This method is only supported for AF systems that use an ADwin device. " & vbNewLine & _
              "Currently, AF System is set to: " & AFSystem & vbNewLine & vbNewLine & _
              "Please check your AF System settings." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysUp_ADwin = False
    Exit Function
                      
        
Set_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Up position." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysUp_ADwin = False
    Exit Function
    
Board_Is_Missing:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Up position." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysUp_ADwin = False
    Exit Function
    
Boot_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Up position." & vbNewLine & vbNewLine & _
              "Unable to boot the ADWIN device." & _
              vbNewLine & "ADWIN Dev # = " & Trim(str(TTLBoard.BoardNum)) & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysUp_ADwin = False
    Exit Function
    
Set_Relay_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the AF/IRM Relays to the Up position." & vbNewLine & vbNewLine & _
              "One or more calls to set digital output byte on the ADWIN device may have failed." & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetAllRelaysUp_ADwin = False
    Exit Function


End Function

Public Function TrySetAllRelaysToDefaultPosition_ADwin() As Boolean

    'Default position is opposite the AF Relay position
    '(If AF's are Relay Up, then Defaults are Relay down and vice versa)
    If modConfig.AFRelays_UseUpPosition Then
    
        TrySetAllRelaysToDefaultPosition_ADwin = TrySetAllRelaysDown_ADwin
        
    Else
    
        TrySetAllRelaysToDefaultPosition_ADwin = TrySetAllRelaysUp_ADwin
        
    End If

End Function

Public Function TrySetRelays_ADwin(ByVal CoilType As coil_type, Optional ByVal do_irm_backfield As Boolean = False) As Boolean

    If modConfig.AFRelays_UseUpPosition Then
    
        TrySetRelays_ADwin = TrySetRelays_ADwin_AFsAreUpPosition(CoilType, do_irm_backfield)
        
    Else
    
        TrySetRelays_ADwin = TrySetRelays_ADwin_AFsAreDownPosition(CoilType, do_irm_backfield)
        
    End If
    
End Function

Public Function TrySetRelays_ADwin_AFsAreUpPosition(ByVal CoilType As coil_type, Optional ByVal do_irm_backfield As Boolean = False) As Boolean

    Dim TTLBoard As Board
    Dim BoardName As String
    
    Dim relay_byte_id As Byte
        
    Set TTLBoard = Nothing
    
On Error GoTo Wrong_AF_System_Error
    
        'This function is designed only for the ADwin system to control the relays
        If AFSystem <> "ADWIN" Then GoTo Wrong_AF_System_Error
                    
On Error GoTo 0

   NOCOMM_MODE = False '??? 4/12/17

    'Check for NOCOMM_MODE
    If NOCOMM_MODE = True Then
    
        TrySetRelays_ADwin_AFsAreUpPosition = True
        Exit Function
        
    End If
    
Dim all_up_byte As Byte

all_up_byte = (2 ^ 6) - 1
    
'Turn on Error Handling
On Error GoTo Set_Board_Error
    
        Select Case CoilType
        
            Case coil_type.Axial
            
                BoardName = AFAxialRelay.BoardName
            
                'Snag the board associated with the AfAxial Relay Channel
                Set TTLBoard = SystemBoards(AFAxialRelay.BoardName)
                relay_byte_id = ((2 ^ (AFAxialRelay.ChanNum)) Mod 256)
    
            Case coil_type.Transverse
        
                BoardName = AFTransRelay.BoardName
            
                'Snag the board associated with the AfTrans Relay Channel
                Set TTLBoard = SystemBoards(AFTransRelay.BoardName)
                relay_byte_id = ((2 ^ (AFTransRelay.ChanNum)) Mod 256)
            
            Case coil_type.IRMAxial
            
                'Snag the board associated with the IRM Relay Channel
                BoardName = IRMRelay.BoardName
                Set TTLBoard = SystemBoards(IRMRelay.BoardName)
            
                If do_irm_backfield Then
                    '??? //4/13/17
                    'Set Transverse high, Axial low, IRM high (Original Implemebtatub)
                    'relay_byte_id = ((2 ^ (IRMRelay.ChanNum)) Mod 256) + ((2 ^ (AFTransRelay.ChanNum)) Mod 256)
                    
                    'Set Transverse low, Axial low, IRM high (New Implemenation)
                    'relay_byte_id = ((2 ^ (IRMRelay.ChanNum)) Mod 256)
                    
                    'Set Transverse High, Axial high, IRM Low (New Implemenation)
                    relay_byte_id = ((2 ^ (IRMRelay.ChanNum)) Mod 256)
                    
                    '??? //4/13/17
                    
                Else
                    'Set Transverse high, Axial low, IRM low
                    relay_byte_id = ((2 ^ (AFTransRelay.ChanNum)) Mod 256)
                
                End If
            
            Case Else
                    
                Err.Raise -1, "frmADWIN_AF.TrySetRelays_ADwin_AFsAreUpPosition", _
                          "Unrecognized or unknown Coil Type given by calling method to this method. " & _
                          "Cannot find ADwin device corresponding to this Coil Type."
        
        End Select
    
    On Error GoTo 0
        
    On Error GoTo Board_Is_Missing
    
    'Error Check again
    If TTLBoard Is Nothing Then GoTo Board_Is_Missing
    
    On Error GoTo 0
    
    On Error GoTo Boot_Board_Error
    
    'Turn off the error-pop-up in the Boot process
    ADWIN.Show_Errors (0)
    
    'Check to make sure the ADWIN board is booted
    If ADWIN.ADWIN_BootBoard(TTLBoard) = False Then GoTo Boot_Board_Error
    
    On Error GoTo 0
    
    On Error GoTo Get_Board_Status_Byte_Error
    
    'Get the current state of the ADwin digital relays
    Dim status_byte As Byte
    
    If Not TTLBoard.TryGetADWIN_DigOutStatusByte(status_byte) Then
    
        'Assume all relay channels are in logic low-state
        status_byte = 0
        
    End If
    
    On Error GoTo 0
    
    On Error GoTo Calculate_Output_Byte_For_ADwin_Error
    
    Dim intermediate_byte As Byte
    Dim final_byte As Byte
    
    Select Case CoilType
    
        Case coil_type.Axial, coil_type.Transverse
            'For AF's, only one relay up at a time
            
            'Intermediate State - Prior State + desired Relay set low
            If (status_byte And relay_byte_id) = relay_byte_id Then
                
                intermediate_byte = status_byte - relay_byte_id
                
            Else
            
                intermediate_byte = status_byte
            
            End If
            
            'Final State - Only Desired Relay set high
            final_byte = relay_byte_id
        
        Case coil_type.IRMAxial, coil_type.IRMTrans
            
            'intermediate - drop all relays
            intermediate_byte = 0
            
            'Final state - set to desired byte
            final_byte = relay_byte_id
            
        Case Else
        
            'Error!
            Err.Raise -1, "frmADWIN_AF.TrySetRelays_ADwin_AFsAreUpPosition", _
                      "Invalid relay position request. The caller asked this method to set the Relays " & _
                      "for an unknown AF or IRM coil type.  " & _
                      "Only Axial and Transverse coil types are currently supported."
    
    End Select

On Error GoTo 0



On Error GoTo Set_Relay_Error

    Dim prior_status As String
    prior_status = frmProgram.sbStatusBar.Panels(2).text
    frmProgram.StatusBar "Setting AF/IRM Relays", 2

    'Set the intermediate relay position
    If intermediate_byte <> final_byte Then
        ADWIN.Set_Digout CLng(intermediate_byte)
    
        'Wait two seconds, do not allow events
        PauseTill_NoEvents timeGetTime() + 2000
    End If
    
    'Set the final relay position
    ADWIN.Set_Digout CLng(final_byte)
        
    PauseTill_NoEvents timeGetTime() + 1000
    
    frmProgram.StatusBar prior_status, 2
    
On Error GoTo 0
    
    TrySetRelays_ADwin_AFsAreUpPosition = True
    Set TTLBoard = Nothing
    Exit Function
        
Dim error_subject As String
Dim error_message As String
        
Wrong_AF_System_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "AF System currently not set to ""ADWIN"". " & vbNewLine & _
              "This method is only supported for AF systems that use an ADwin device. " & vbNewLine & _
              "Currently, AF System is set to: " & AFSystem & vbNewLine & vbNewLine & _
              "Please check your AF System settings." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreUpPosition = False
    Exit Function
                      
        
Set_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreUpPosition = False
    Exit Function
    
Board_Is_Missing:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreUpPosition = False
    Exit Function
    
Boot_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "Unable to boot the ADWIN device." & _
              vbNewLine & "ADWIN Dev # = " & Trim(str(TTLBoard.BoardNum)) & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreUpPosition = False
    Exit Function
    
Get_Board_Status_Byte_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & _
              "Unable to get the current status of the AF and IRM Relays from the ADWIN device." & _
              vbNewLine & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreUpPosition = False
    Exit Function
    
Calculate_Output_Byte_For_ADwin_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & _
              "Error was generated by the code used to determined what digital output byte " & _
              "value to set to the ADWIN device." & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreUpPosition = False
    Exit Function
    
Set_Relay_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & _
              "One or more calls to set digital output byte on the ADWIN device may have failed." & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreUpPosition = False
    Exit Function
    
End Function

Public Function TrySetRelays_ADwin_AFsAreDownPosition(ByVal CoilType As coil_type, Optional ByVal do_irm_backfield As Boolean = False) As Boolean

    Dim TTLBoard As Board
    Dim BoardName As String
    
    Dim relay_byte_id As Byte
        
    Set TTLBoard = Nothing
    
On Error GoTo Wrong_AF_System_Error
    
        'This function is designed only for the ADwin system to control the relays
        If AFSystem <> "ADWIN" Then GoTo Wrong_AF_System_Error
                    
On Error GoTo 0

    'Check for NOCOMM_MODE
    NOCOMM_MODE = False '??? 5/5/17
    
    If NOCOMM_MODE = True Then
    
        TrySetRelays_ADwin_AFsAreDownPosition = True
        Exit Function
        
    End If
    
Dim all_up_byte As Byte

all_up_byte = (2 ^ 6) - 1
    
'Turn on Error Handling
On Error GoTo Set_Board_Error
    
        Select Case CoilType
        
            Case coil_type.Axial
            
                BoardName = AFAxialRelay.BoardName
            
                'Snag the board associated with the AfAxial Relay Channel
                Set TTLBoard = SystemBoards(AFAxialRelay.BoardName)
                relay_byte_id = all_up_byte - ((2 ^ (AFAxialRelay.ChanNum)) Mod 256)
    
            Case coil_type.Transverse
        
                BoardName = AFTransRelay.BoardName
            
                'Snag the board associated with the AfTrans Relay Channel
                Set TTLBoard = SystemBoards(AFTransRelay.BoardName)
                relay_byte_id = all_up_byte - ((2 ^ (AFTransRelay.ChanNum)) Mod 256)
            
            Case coil_type.IRMAxial
            
                'Snag the board associated with the IRM Relay Channel
                BoardName = IRMRelay.BoardName
                Set TTLBoard = SystemBoards(IRMRelay.BoardName)
            
                If do_irm_backfield Then
                    
                    'Set Transverse high, Axial high, IRM high
                    'relay_byte_id = all_up_byte '??? Original Implementation //5/5/17
                    
                    relay_byte_id = all_up_byte - ((2 ^ (IRMRelay.ChanNum)) Mod 256)
                    
                Else
                    'Set IRM Transverse high, Axial High, IRM low
                    relay_byte_id = all_up_byte - ((2 ^ (IRMRelay.ChanNum)) Mod 256)
                
                End If
            
            Case Else
                    
                Err.Raise -1, "frmADWIN_AF.TrySetRelays_ADwin_AFsAreDownPosition", _
                          "Unrecognized or unknown Coil Type given by calling method to this method. " & _
                          "Cannot find ADwin device corresponding to this Coil Type."
        
        End Select
    
    On Error GoTo 0
        
    On Error GoTo Board_Is_Missing
    
    'Error Check again
    If TTLBoard Is Nothing Then GoTo Board_Is_Missing
    
    On Error GoTo 0
    
    On Error GoTo Boot_Board_Error
    
    'Turn off the error-pop-up in the Boot process
    ADWIN.Show_Errors (0)
    
    'Check to make sure the ADWIN board is booted
    If ADWIN.ADWIN_BootBoard(TTLBoard) = False Then GoTo Boot_Board_Error
    
    On Error GoTo 0
    
    On Error GoTo Get_Board_Status_Byte_Error
    
    'Get the current state of the ADwin digital relays
    Dim status_byte As Byte
    
    If Not TTLBoard.TryGetADWIN_DigOutStatusByte(status_byte) Then
    
        'Assume all relay channels are in logic high-state
        status_byte = all_up_byte
        
    End If
    
    On Error GoTo 0
    
    On Error GoTo Calculate_Output_Byte_For_ADwin_Error
    
    Dim intermediate_byte As Byte
    Dim final_byte As Byte
    
    Select Case CoilType
    
        Case coil_type.Axial, coil_type.Transverse
            'For AF's, only one relay up at a time
            
            'Intermediate State - Prior State + desired Relay set high
            intermediate_byte = all_up_byte
            
            'Final State - Only Desired Relay set high
            final_byte = relay_byte_id
        
        Case coil_type.IRMAxial, coil_type.IRMTrans
            
            'intermediate - raise all non-IRM relays
            intermediate_byte = all_up_byte
            
            'Final state - set to desired byte
            final_byte = relay_byte_id
            
        Case Else
        
            'Error!
            Err.Raise -1, "frmADWIN_AF.TrySetRelays_ADwin_AFsAreDownPosition", _
                      "Invalid relay position request. The caller asked this method to set the Relays " & _
                      "for an unknown AF or IRM coil type.  " & _
                      "Only Axial and Transverse coil types are currently supported."
    
    End Select

On Error GoTo 0



On Error GoTo Set_Relay_Error

    Dim prior_status As String
    prior_status = frmProgram.sbStatusBar.Panels(2).text
    frmProgram.StatusBar "Setting AF/IRM Relays", 2

    'Set the intermediate relay position
    If intermediate_byte <> final_byte Then
        ADWIN.Set_Digout CLng(intermediate_byte)
    
        'Wait two seconds, do not allow events
        PauseTill_NoEvents timeGetTime() + 2000
    End If
    
    'Set the final relay position
    ADWIN.Set_Digout CLng(final_byte)
        
    PauseTill_NoEvents timeGetTime() + 1000
    
    frmProgram.StatusBar prior_status, 2
    
On Error GoTo 0
    
    TrySetRelays_ADwin_AFsAreDownPosition = True
    Set TTLBoard = Nothing
    Exit Function
        
Dim error_subject As String
Dim error_message As String
        
Wrong_AF_System_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "AF System currently not set to ""ADWIN"". " & vbNewLine & _
              "This method is only supported for AF systems that use an ADwin device. " & vbNewLine & _
              "Currently, AF System is set to: " & AFSystem & vbNewLine & vbNewLine & _
              "Please check your AF System settings." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreDownPosition = False
    Exit Function
                      
        
Set_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
    
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreDownPosition = False
    Exit Function
    
Board_Is_Missing:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "Board: """ & BoardName & """ " & vbNewLine & _
              "is missing from the System Boards collection." & vbNewLine & vbNewLine & _
              "Check your system settings and the .INI file [Boards] & [Channels] " & _
              "sections." & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreDownPosition = False
    Exit Function
    
Boot_Board_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & vbNewLine & _
              "Unable to boot the ADWIN device." & _
              vbNewLine & "ADWIN Dev # = " & Trim(str(TTLBoard.BoardNum)) & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreDownPosition = False
    Exit Function
    
Get_Board_Status_Byte_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & _
              "Unable to get the current status of the AF and IRM Relays from the ADWIN device." & _
              vbNewLine & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreDownPosition = False
    Exit Function
    
Calculate_Output_Byte_For_ADwin_Error:
    
    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & _
              "Error was generated by the code used to determined what digital output byte " & _
              "value to set to the ADWIN device." & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreDownPosition = False
    Exit Function
    
Set_Relay_Error:

    'Set Code Red
    modFlow.Flow_Pause
    modStatusCode.SetCodeLevel CodeRed
    
    error_subject = "AF Relay Error!"
    error_message = _
              "Error occurred while trying to Set the Relay position for an ADwin AF Ramp." & vbNewLine & _
              "One or more calls to set digital output byte on the ADWIN device may have failed." & vbNewLine & _
              "Error Detail: " & Err.Description & vbNewLine & vbNewLine & _
              "Code Execution has been paused."
              
    frmSendMail.MailNotification error_subject, _
                                 error_message, _
                                 CodeRed
                                 
    MsgBox error_message, vbCritical, error_subject
                  
    modStatusCode.SetCodeLevel modStatusCode.StatusCodeColorLevelPrior
                  
    TrySetRelays_ADwin_AFsAreDownPosition = False
    Exit Function
    
End Function


Private Sub SetPeakValues()

    'Now, depending on the coil system used, need to translate the active field
    'setting (Peak Field or Monitor Voltage) into the other two fields (Ramp voltage needs
    'to be set as well).
    'Check if the Ramp is unmonitored
    If Me.chkClippingTest.value = Unchecked Then
    
        'This is a monitored ramp, the peak ramp value needs
        'to be determined from the Peak MOnitor Voltage
    
        'If this is a calibrated ramp
        'Need to get the needed Monitor peak voltage
        'from the Peak Field value
        
        'Check to see if the Axial coil is active
        If ActiveCoilSystem = AxialCoilSystem Then
        
            'Check to see if this AF coil has been calibrated
            'and if it needs to be
            If AFAxialCalDone = False And _
               optCalRamp(0).value = True _
            Then
            
                'Tell user they need to do a calibration
                'on the AF Axial coil
                UserResp = MsgBox("AF Axial Coil has not been calibrated yet. The current AF " & _
                                  "ramp process has been terminated." & vbNewLine & vbNewLine & _
                                  "Would you like to calibrate the AF Axial Coil right now?", _
                                  vbYesNo, _
                                  "Ooops!")
                                  
                If UserResp = vbYes Then
                
                    frmCalibrateCoils.InAFMode = True
                
                    'Load the AF calibration form
                    Load frmCalibrateCoils
                                        
                    'Open the form
                    frmCalibrateCoils.Show
                    
                    'Pause the program flow
                    Flow_Pause
                
                    'Wait for the flow to be unpaused
                    modFlow.Flow_WaitForUnpaused
                    
                End If
        
                Exit Sub
            
            End If
                        
            'Now check to see if the user wants a calibrated, monitored ramp
            If optCalRamp(0).value = True Then
                    
                'We know the coil system is calibrated, we can get the Monitor voltage
                'matching the Peak Field set by the user
                Me.txtMonitorTrigVolt = Format(frmAF_2G.FindXCalibValue(val(Me.txtPeakField), _
                                                                          ActiveCoilSystem), "#0.0#####")
                                                                          
                'Check to make sure the monitor voltage is below the max monitor voltage for
                'the Axial coil system
                If val(Me.txtMonitorTrigVolt) > modConfig.AfAxialMonMax Then
                
                    Me.txtMonitorTrigVolt = Format(modConfig.AfAxialMonMax, "#0.0#####")
                    
                End If
                
            Else
            
                'Check to make sure the monitor voltage is below the max monitor voltage for
                'the Axial coil system
                If val(Me.txtMonitorTrigVolt) > modConfig.AfAxialMonMax Then
                
                    Me.txtMonitorTrigVolt = Format(modConfig.AfAxialMonMax, "#0.0#####")
                    
                End If
            
                'The user wants an uncalibrated ramp using the Peak Monitor Voltage value
                'If this coil has been calibrated, then update the Peak Field text-box
                If AFAxialCalDone = True Then
                
                    Me.txtPeakField = Format(modAF_DAQ.FindFieldFromVolts( _
                                                            val(Me.txtMonitorTrigVolt), _
                                                            ActiveCoilSystem), _
                                             "#0.0##")
                End If
            
            End If
            
        ElseIf ActiveCoilSystem = TransverseCoilSystem Then
        
            'Check to see if this AF coil has been calibrated
            'and if it needs to be
            If AFTransCalDone = False And _
               optCalRamp(0).value = True _
            Then
            
                'Tell user they need to do a calibration
                'on the AF Transverse coil
                UserResp = MsgBox("AF Transverse Coil has not been calibrated yet. The current AF " & _
                                  "ramp cycle has been terminated." & vbNewLine & vbNewLine & _
                                  "Would you like to calibrate the AF Transverse Coil right now?", _
                                  vbYesNo, _
                                  "Ooops!")
                                  
                If UserResp = vbYes Then
                
                    frmCalibrateCoils.InAFMode = True
                
                    'Load the AF calibration form
                    Load frmCalibrateCoils
                    
                    'Open the form
                    frmCalibrateCoils.Show
                    
                    'Pause the program flow
                    Flow_Pause
                
                    'Wait for the flow to be unpaused
                    modFlow.Flow_WaitForUnpaused
                                        
                End If
        
                Exit Sub
            
            End If
            
            'Now check to see if the user wants a calibrated, monitored ramp
            If optCalRamp(0).value = True Then
                    
                'We know the coil system is calibrated, we can get the Monitor voltage
                'matching the Peak Field set by the user
                Me.txtMonitorTrigVolt = Format(frmAF_2G.FindXCalibValue(val(Me.txtPeakField), _
                                                                          TransverseCoilSystem), _
                                               "#0.0#####")
                                                                          
                'Check to make sure the monitor voltage is below the max monitor voltage for
                'the Transverse coil system
                If val(Me.txtMonitorTrigVolt) > modConfig.AfTransMonMax Then
                
                    Me.txtMonitorTrigVolt = Format(modConfig.AfTransMonMax, "#0.0#####")
                    
                End If
                
            Else
            
                'Check to make sure the monitor voltage is below the max monitor voltage for
                'the Transverse coil system
                If val(Me.txtMonitorTrigVolt) > modConfig.AfTransMonMax Then
                
                    Me.txtMonitorTrigVolt = Format(modConfig.AfTransMonMax, "#0.0#####")
                    
                End If
            
                'The user wants an uncalibrated ramp using the Peak Monitor Voltage value
                'If this coil has been calibrated, then update the Peak Field text-box
                If AFTransCalDone = True Then
                
                    Me.txtPeakField = Format(modAF_DAQ.FindFieldFromVolts( _
                                                            val(Me.txtMonitorTrigVolt), _
                                                            TransverseCoilSystem), _
                                             "#0.0##")
                End If
            
            End If
            
        End If
                
        'Update the Ramp Peak Voltage
        'Depends on how high the monitor voltage relative to the maximum monitor voltage
        If ActiveCoilSystem = AxialCoilSystem Then
        
            Me.txtRampPeakVoltage = Format(val(Me.txtMonitorTrigVolt) _
                                           / modConfig.AfAxialMonMax _
                                           * modConfig.AfAxialRampMax, _
                                           "#0.0#####")
                                             
        ElseIf ActiveCoilSystem = TransverseCoilSystem Then
        
            Me.txtRampPeakVoltage = Format(val(Me.txtMonitorTrigVolt) _
                                           / modConfig.AfTransMonMax _
                                           * modConfig.AfTransRampMax / 2.2, _
                                           "#0.0#####")
                                                                            
        End If
                                                                            
        'Make sure the peak ramp voltage is within bounds
        If ActiveCoilSystem = AxialCoilSystem And _
           val(Me.txtRampPeakVoltage) > modConfig.AfAxialRampMax _
        Then
        
            Me.txtRampPeakVoltage = Format(modConfig.AfAxialRampMax, "#0.0#####")
            
        ElseIf ActiveCoilSystem = TransverseCoilSystem And _
               val(Me.txtRampPeakVoltage) > modConfig.AfTransRampMax _
        Then
        
            Me.txtRampPeakVoltage = Format(modConfig.AfTransRampMax, "#0.0#####")
        
        End If
                                                                                                                                   
    End If
    
End Sub
'??? 5/1/17
Private Sub SWITCH_BUTTON_Click()

  Dim result As Boolean
     'Default position is opposite the AF Relay position
    '(If AF's are Relay Up, then Defaults are Relay down and vice versa)
    If optCoil(0).value = True Then
         result = TrySetRelays_ADwin(Axial, True)        ' true = backfield
    ElseIf optCoil(1).value = True Then
         result = TrySetRelays_ADwin(Transverse, True)   ' true = backfield
    ElseIf optCoil(2).value = True Then
         result = TrySetRelays_ADwin(IRMAxial, True)     ' true = backfield
    End If
    
End Sub
'??? /5/1/17

Private Sub txtFreq_Change()

    If val(txtFreq) <> 0 Then
        Me.txtRampPeakDuration = Trim(str(1 / val(txtFreq) * 100000))
    End If
    
End Sub

Private Sub txtRampDownSlope_Change()

    Dim TempD As Double
    Dim TempS As String
    
    'Store the value of the Ramp Up slope to local var.
    TempD = val(Me.txtRampDownSlope)

    'Check to make slope is greater than zero
    If TempD <= 0 Then Exit Sub

    'Need to adjust the duration label next to the RampDown Slope
    TempS = Trim(str(val(Me.txtRampPeakVoltage) / TempD * 1000))
    Me.lblRampDownDuration.Caption = PadLeft(TempS, 6)
                                     
    'Need to adjust the total duration label
    'Add the ramp up and ramp down durations + time at peak
    'and the extra 200 ms the code adds to make sure the process
    'has indeed finished
    TempD = val(Me.lblRampDownDuration.Caption) + _
            val(Me.lblRampUpDuration.Caption) + _
            val(Me.txtRampPeakDuration) + _
            200
            
    'Update the Total duration label
    Me.lblTotalRampDuration = PadLeft(Trim(str(CLng(TempD))), 7)
    
End Sub

Private Sub txtRampDownSlope_LostFocus()

    Dim TempD As Double
    
    'Store the value of the Ramp Up slope to local var.
    TempD = val(Me.txtRampDownSlope)

    'Check to make slope is greater than zero
    If TempD < 0 Then
    
        'Flip the sign
        Me.txtRampDownSlope = Trim(str(-1 * TempD))
        
        'Activate the value change event
        'to update the duration label
        txtRampDownSlope_Change
        
    End If

End Sub

Private Sub txtRampPeakVoltage_Change()

    'If any change is made to this value need to call the txtRampUpSlope
    'and txtRampDownSlope Change events
    
    'Calculate the Ramp Up and Down slopes to use
    txtRampUpSlope = Trim(str(GetUpSlope(val(Me.txtRampPeakVoltage))))
    txtRampDownSlope = Trim(str(GetDownSlope(val(Me.txtRampPeakVoltage))))
        
    txtRampUpSlope_Change
    txtRampDownSlope_Change

End Sub

Private Sub txtRampUpSlope_Change()

    Dim TempD As Double
    Dim TempS As String
    
    'Store the value of the Ramp Up slope to local var.
    TempD = val(Me.txtRampUpSlope)

    'Check to make slope is greater than zero
    If TempD <= 0 Then Exit Sub

    'Need to adjust the duration label next to the RampUp Slope
    TempS = Trim(str(val(Me.txtRampPeakVoltage) / TempD * 1000))
    Me.lblRampUpDuration.Caption = PadLeft(TempS, 6)
                                     
    'Need to adjust the total duration label
    'Add the ramp up and ramp down durations + time at peak
    'and the extra 200 ms the code adds to make sure the process
    'has indeed finished
    TempD = val(Me.lblRampUpDuration.Caption) + _
            val(Me.lblRampDownDuration.Caption) + _
            val(Me.txtRampPeakDuration) + _
            200
            
    'Update the Total duration label
    Me.lblTotalRampDuration = PadLeft(Trim(str(CLng(TempD))), 7)
    
    
End Sub

Private Sub txtRampUpSlope_LostFocus()

    Dim TempD As Double
    
    'Store the value of the Ramp Up slope to local var.
    TempD = val(Me.txtRampUpSlope)

    'Check to make slope is greater than zero
    If TempD < 0 Then
    
        'Flip the sign
        Me.txtRampUpSlope = Trim(str(-1 * TempD))
        
        'Activate the value change event
        'to update the duration label
        txtRampUpSlope_Change
        
    End If

End Sub

