VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRockmagRoutine 
   Caption         =   "Set Rock Mag Routine"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmRockmagRoutine.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   9600
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmPresets 
      Caption         =   "Presets"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.CheckBox chkTheWorksMeasSusc 
         Caption         =   "Measure Susceptibility"
         Height          =   372
         Left            =   5280
         TabIndex        =   67
         Top             =   1150
         Width           =   2055
      End
      Begin VB.CommandButton cmdHawaiianStd 
         Caption         =   "Hawaiian Standard AF (25, 50, 100, 200, 400, 800)"
         Height          =   372
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdRockmagEverything 
         Caption         =   "Rockmag ""the Works"""
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox chkRMAllNRM 
         Caption         =   "Measure and AF demagnetize NRM"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkRMAllNRM3AxisAF 
         Caption         =   "along all three axes"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox chkRMAllWithRRM 
         Caption         =   "RRM (rps):"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtRMAllRRMrpsStep 
         Height          =   288
         Left            =   2280
         TabIndex        =   6
         Top             =   1560
         Width           =   732
      End
      Begin VB.TextBox txtRMAllRRMrpsMax 
         Height          =   288
         Left            =   3360
         TabIndex        =   7
         Top             =   1560
         Width           =   732
      End
      Begin VB.TextBox txtRMAllRRMAFField 
         Height          =   288
         Left            =   4680
         TabIndex        =   8
         Top             =   1560
         Width           =   732
      End
      Begin VB.CheckBox chkRMAllRRMdoNegative 
         Caption         =   "and negative rotations"
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox chkRMAllARM 
         Caption         =   "ARM Step Size (G):"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtRMAllARMStepSize 
         Height          =   288
         Left            =   2280
         TabIndex        =   11
         Top             =   1920
         Width           =   732
      End
      Begin VB.TextBox txtRMAllARMStepMax 
         Height          =   288
         Left            =   3360
         TabIndex        =   12
         Top             =   1920
         Width           =   732
      End
      Begin VB.TextBox txtRMAllAFFieldForARM 
         Height          =   288
         Left            =   4680
         TabIndex        =   13
         Top             =   1920
         Width           =   732
      End
      Begin VB.CheckBox chkRMAllIRM 
         Caption         =   "AF/IRM log step (G):"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtRMAllLogFactor 
         Height          =   288
         Left            =   2280
         TabIndex        =   15
         Top             =   2280
         Width           =   732
      End
      Begin VB.TextBox txtRMAllMinStepSize 
         Height          =   288
         Left            =   4680
         TabIndex        =   16
         Top             =   2280
         Width           =   732
      End
      Begin VB.TextBox txtRMAllAFMax 
         Height          =   288
         Left            =   6480
         TabIndex        =   17
         Top             =   2280
         Width           =   732
      End
      Begin VB.TextBox txtRMAllIRMMax 
         Height          =   288
         Left            =   8280
         TabIndex        =   18
         Top             =   2280
         Width           =   732
      End
      Begin VB.CheckBox chkRMAllBackfieldDemag 
         Caption         =   "DC Demag via backfield IRM"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "to"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "in AF "
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "to"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "in AF "
         Height          =   255
         Left            =   4200
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Min. step size (G):"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "AF Max (G):"
         Height          =   255
         Left            =   5520
         TabIndex        =   25
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "IRM Max (G):"
         Height          =   255
         Left            =   7320
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.Frame frameSetSteps 
      Caption         =   "Set Steps"
      Height          =   2175
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   9255
      Begin VB.TextBox txtStepMin 
         Height          =   288
         Left            =   2040
         TabIndex        =   65
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox txtStepSize 
         Height          =   288
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox txtStepMax 
         Height          =   288
         Left            =   3240
         TabIndex        =   29
         Top             =   360
         Width           =   732
      End
      Begin VB.ComboBox cmbStepSeq 
         Height          =   315
         Left            =   4920
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbStepSeqScale 
         Height          =   315
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdaddStepSeq 
         Caption         =   "Add"
         Height          =   372
         Left            =   7920
         TabIndex        =   32
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox txtARMSteps 
         Height          =   288
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   732
      End
      Begin VB.TextBox txtARMBiasMax 
         Height          =   288
         Left            =   2880
         TabIndex        =   34
         Top             =   840
         Width           =   732
      End
      Begin VB.TextBox txtAFfieldForARM 
         Height          =   288
         Left            =   4920
         TabIndex        =   35
         Top             =   840
         Width           =   732
      End
      Begin VB.ComboBox cmbARMStepSeqScale 
         Height          =   315
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   36
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdAddARMStepSeq 
         Caption         =   "Add"
         Height          =   372
         Left            =   7920
         TabIndex        =   37
         Top             =   960
         Width           =   1092
      End
      Begin VB.ComboBox cmbStepType 
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtLevel 
         Height          =   288
         Left            =   1440
         TabIndex        =   39
         Top             =   1560
         Width           =   732
      End
      Begin VB.TextBox txtBiasField 
         Height          =   288
         Left            =   2280
         TabIndex        =   40
         Top             =   1560
         Width           =   732
      End
      Begin VB.TextBox txtSpinSpeed 
         Height          =   288
         Left            =   3120
         TabIndex        =   41
         Top             =   1560
         Width           =   732
      End
      Begin VB.TextBox txtHoldTime 
         Height          =   288
         Left            =   3960
         TabIndex        =   42
         Top             =   1560
         Width           =   732
      End
      Begin VB.CheckBox chkMeasure 
         Caption         =   "Measure"
         Height          =   372
         Left            =   4800
         TabIndex        =   43
         Top             =   1560
         Width           =   950
      End
      Begin VB.CheckBox chkSusceptibility 
         Caption         =   "Susceptibility"
         Height          =   372
         Left            =   5760
         TabIndex        =   44
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtRemarks 
         Height          =   288
         Left            =   7080
         TabIndex        =   45
         Top             =   1560
         Width           =   732
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   372
         Left            =   7920
         TabIndex        =   46
         Top             =   1560
         Width           =   1092
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Height          =   375
         Left            =   7920
         TabIndex        =   47
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "to"
         Height          =   255
         Left            =   3000
         TabIndex        =   66
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "G steps from"
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "G of type"
         Height          =   255
         Left            =   4080
         TabIndex        =   49
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "G ARM bias steps up to"
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "G in AF field of"
         Height          =   255
         Left            =   3720
         TabIndex        =   51
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "G"
         Height          =   255
         Left            =   5760
         TabIndex        =   52
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Step Type"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Level"
         Height          =   255
         Left            =   1440
         TabIndex        =   54
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Bias (G)"
         Height          =   255
         Left            =   2280
         TabIndex        =   55
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Spin (rps)"
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Hold (s)"
         Height          =   255
         Left            =   3960
         TabIndex        =   57
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   7080
         TabIndex        =   58
         Top             =   1320
         Width           =   735
      End
   End
   Begin ComctlLib.ListView lvwSteps 
      Height          =   2175
      Left            =   120
      TabIndex        =   59
      Top             =   5640
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3836
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   372
      Left            =   480
      TabIndex        =   60
      Top             =   8040
      Width           =   1092
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   372
      Left            =   1920
      TabIndex        =   61
      Top             =   8040
      Width           =   1092
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import from .RMG"
      Height          =   375
      Left            =   3840
      TabIndex        =   62
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save to .RMG"
      Height          =   375
      Left            =   6240
      TabIndex        =   63
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   8160
      TabIndex        =   64
      Top             =   8040
      Width           =   1215
   End
End
Attribute VB_Name = "frmRockmagRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------'
'
'   Major Code Mod
'   August 14, 2010
'   Isaac Hilburn
'
'   Summary:    Added in filter in Form Load and Form Activate event handlers to disable/enable controls based on
'               enabled/disabled state of the Rockmag modules that those controls are used to setup a rockmag run for.
'               For instance, if the ARM module is switched off, all of the controls allowing the user to setup
'               an ARM run would be switched off.
'
'               Also, in the middle frame, I added a new text-box: txtStepMin and changed the code in cmdAdd_Click
'               to calculate the steps from a user inputed minimum value instead of zero.  (This lack of functionality
'               has been bugging me for years!)
'
'------------------------------------------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------------------------------------------'

Dim ActiveSAMFile As Integer
Dim SequenceReady As Boolean
Public rmStepList As RockmagSteps
Private rmStepNumber As Integer

Private Sub chkRMAllNRM_Click()
    If chkRMAllNRM = Checked Then
        chkRMAllNRM3AxisAF.Enabled = True
    Else
        chkRMAllNRM3AxisAF.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim i As Double
    If cmbStepType.text = "NRM" Then ' (August 2007 L Carporzen) NRM possible in RockMag
    rmStepList.Add cmbStepType.text, 0, 0, 0, _
        0, (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    ElseIf cmbStepType.text = "X" Then ' (June 2009 L Carporzen) Susceptibility only
    rmStepList.Add cmbStepType.text, "0", "0", "0", _
        "0", Unchecked, Checked, txtRemarks
    ElseIf cmbStepType.text = "GRM AF" Then ' (Sept 2008 L Carporzen) Uniaxial AF: Gyromagnetic remanent magnetization demag from Stephenson 1993
    cmbStepType.text = "UAFX2"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "UAFZ1"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "UAFX1"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "UAFX2"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "UAFZ1"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "GRM AF"
    ElseIf cmbStepType.text = "UAF" Then ' (March 2008 L Carporzen) Uniaxial AF: measure the sample after each axis demag
    cmbStepType.text = "UAFX1"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "UAFX2"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "UAFZ1"
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    cmbStepType.text = "UAF"
    Else
    rmStepList.Add cmbStepType.text, val(txtLevel), val(txtBiasField), val(txtSpinSpeed), _
        val(txtHoldTime), (chkMeasure.value = Checked), (chksusceptibility.value = Checked), txtRemarks
    End If
    refreshListDisplay
End Sub

Private Sub cmdaddARMStepSeq_Click()
    Dim StepSize As Double
    Dim StepMax As Double
    Dim NumSteps As Integer
    Dim curItem As ListItem
    Dim i As Double
    StepSize = val(txtARMSteps)
    StepMax = val(txtARMBiasMax)
    If EnableARM = False Then Exit Sub
    If StepSize = 0 Or StepSize > StepMax Then Exit Sub
    If cmbStepSeqScale = "Log" Then
        rmStepList.Add cmbStepSeq.text, 0
        If StepSize = 1 Then Exit Sub
        NumSteps = Log(StepMax) / Log(StepSize)
        If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)

        For i = 1 To Log(StepMax) / Log(StepSize)
            rmStepList.Add cmbStepSeq.text, val(txtAFfieldForARM), StepSize ^ i
        Next i
    Else
        If StepMax Mod StepSize <> 0 Then StepMax = StepMax - StepMax Mod StepSize
        For i = 0 To StepMax Step StepSize
            rmStepList.Add "ARM", val(txtAFfieldForARM), i
        Next i
    End If
    refreshListDisplay
End Sub

Private Sub cmdaddStepSeq_Click()
    Dim StepSize As Double
    Dim StepMax As Double
    Dim StepMin As Double
    Dim NumSteps As Integer
    Dim curItem As ListItem
    Dim CurStep As Double
    Dim i As Double
    
    StepSize = val(txtStepSize)
    StepMax = val(txtStepMax)
    
    '(August 2010 - I Hilburn) Added in StepMin to step calculation
    StepMin = val(txtStepMin)
    
    '(August 2010 - I Hilburn) Added in StepMin to step calculation
    If StepSize = 0 Or _
       StepSize > StepMax Or _
       StepMin > StepMax _
    Then Exit Sub
    
    If cmbStepSeqScale = "Log" Then
        
        '(August 2010 - I Hilburn) Added in StepMin to step calculation
        rmStepList.Add cmbStepSeq.text, _
                       StepMin, _
                       val(txtBiasField), _
                       val(txtSpinSpeed), _
                       val(txtHoldTime), _
                       (chkMeasure.value = Checked), _
                       (chksusceptibility.value = Checked), _
                       txtRemarks
                       
        If StepSize = 1 Then Exit Sub
        
        '(August 2010 - I Hilburn) Added in StepMin to step calculation
        NumSteps = Log(StepMax - StepMin) / Log(StepSize)
        
        If NumSteps - CInt(NumSteps) > 0.5 Then
        
            NumSteps = CInt(NumSteps) + 1
            
        Else
        
            NumSteps = CInt(NumSteps)
            
        End If
        
        For i = 1 To NumSteps
        
            '(August 2010 - I Hilburn) Added in StepMin to step calculation
            CurStep = StepMin + CInt(StepSize ^ i)
            
            rmStepList.Add cmbStepSeq.text, _
                           CurStep, _
                           val(txtBiasField), _
                           val(txtSpinSpeed), _
                           val(txtHoldTime), _
                           (chkMeasure.value = Checked), _
                           (chksusceptibility.value = Checked), _
                           txtRemarks
                           
        Next i
        
    Else
    
        '(August 2010 - I Hilburn) Added in StepMin to step calculation
        If (StepMax - StepMin) Mod StepSize <> 0 Then
        
            StepMax = StepMax - (StepMax - StepMin) Mod StepSize
            
        End If
        
        '(August 2010 - I Hilburn) Added in StepMin to step calculation
'        rmStepList.add cmbStepSeq.text, _
'                       StepMin, _
'                       val(txtBiasField), _
'                       val(txtSpinSpeed), _
'                       val(txtHoldTime), _
'                       (chkMeasure.Value = Checked), _
'                       (chksusceptibility.Value = Checked), _
'                       txtRemarks
                       
        For i = StepMin To StepMax Step StepSize
            
            '(August 2010 - I Hilburn) Added in StepMin to step calculation
            rmStepList.Add cmbStepSeq.text, _
                           CInt(i), _
                           val(txtBiasField), _
                           val(txtSpinSpeed), _
                           val(txtHoldTime), _
                           (chkMeasure.value = Checked), _
                           (chksusceptibility.value = Checked), _
                           txtRemarks
                           
        Next i
        
    End If
    
    refreshListDisplay
    
End Sub

Private Sub cmdClear_Click()
    Set rmStepList = Nothing
    Set rmStepList = New RockmagSteps
    refreshListDisplay
End Sub

Private Sub cmdDelete_Click()
    Dim targetItem As ListItem
    If lvwSteps.SelectedItem.index > 0 Then
        For Each targetItem In lvwSteps.ListItems
            If targetItem.Selected Then
                rmStepList.Remove targetItem.key
            End If
        Next targetItem
    Else
        cmdDelete.Enabled = False
    End If
    refreshListDisplay
End Sub

Private Sub cmdHawaiianStd_Click()
    Dim i As Integer
    Set rmStepList = Nothing
    Set rmStepList = New RockmagSteps
    rmStepList.Add "AF", 0
    For i = 1 To 6
        rmStepList.Add "AF", 25 * 2 ^ i
    Next i
    refreshListDisplay
End Sub

Private Sub cmdImport_Click()
    Dim sfilen  As String           ' Path + filename of selected file
    Dim dirname As String           ' Path of selected file
    Dim fname   As String           ' Filename of selected file
    Dim ind     As Variant          ' index of first character of filename
    ' Initialize the dialog box for Sample File Open
    dlgCommonDialog.FILTER = "Rockmag file (*.rmg)|*.rmg|All files (*.*)|*.*"
    dlgCommonDialog.flags = cdlOFNFileMustExist
    dlgCommonDialog.DialogTitle = "Open RMG File..."
    If FileExists(Prog_DefaultPath) Then
            dlgCommonDialog.InitDir = Prog_DefaultPath
        Else
            dlgCommonDialog.InitDir = "\"
        End If
    dlgCommonDialog.ShowOpen
    ' ----- Start parsing the filename -----
    ' Parse the file name
    sfilen = dlgCommonDialog.filename
    If LenB(sfilen) = 0 Then Exit Sub          ' If we don't have a filename
                                               ' then don't processs it.
    ImportRMGRoutine (sfilen)
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    SequenceReady = True
    saveDefaults
    Me.Hide
End Sub

Private Sub cmdReplace_Click()
' (September 2007 L Carporzen) Actualize the line in the RockMag list
    With rmStepList.Item(rmStepNumber)
      .StepType = cmbStepType.text
      If cmbStepType.text = "NRM" Then
        .Level = 0
        .BiasField = 0
        .SpinSpeed = 0
        .HoldTime = 0
      Else
        .Level = val(txtLevel)
        .BiasField = val(txtBiasField)
        .SpinSpeed = val(txtSpinSpeed)
        .HoldTime = val(txtHoldTime)
      End If
      If chkMeasure.value = Checked Then .Measure = True Else .Measure = False
      If chksusceptibility.value = Checked Then .MeasureSusceptibility = True Else .MeasureSusceptibility = False
      .Remarks = txtRemarks ' (November 2007 L Carporzen) Remarks column in RMG
    End With
    refreshListDisplay
    cmdReplace.Visible = False  ' Put back the button "Add"
    cmdAdd.Visible = True  ' Put back the button "Add"
End Sub

Private Sub cmdRockmagEverything_Click()
    Dim StepSize As Double
    Dim stepsizeAF As Double
    Dim StepMax As Long
    Dim stepmaxAF As Long
    Dim stepmaxIRM As Long
    Dim minStepSize As Long
    Dim lastStep As Long
    Dim CurStep As Long
    Dim NumSteps As Double
    Dim field As Long
    Dim i As Double
    'Set rmStepList = Nothing
    'Set rmStepList = New RockmagSteps
    stepsizeAF = val(txtRMAllLogFactor)
    stepmaxAF = val(txtRMAllAFMax)
    If stepmaxAF > AfAxialMax Then stepmaxAF = AfAxialMax
    stepmaxIRM = val(txtRMAllIRMMax)
    minStepSize = val(txtRMAllMinStepSize)
    If stepsizeAF = 0 Or stepsizeAF = 1 Or stepsizeAF > stepmaxAF Or stepsizeAF > stepmaxIRM Then Exit Sub
    If chkRMAllNRM And EnableAF Then
            rmStepList.Add "NRM", MeasureSusceptibility:=(chkTheWorksMeasSusc.value = Checked)
            If chkRMAllNRM3AxisAF = Unchecked Then
                If stepmaxAF > AfAxialMax Then stepmaxAF = AfAxialMax
                NumSteps = Log(stepmaxAF) / Log(stepsizeAF)
                If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)
                lastStep = 0
                For i = 1 To NumSteps
                    CurStep = Int(stepsizeAF ^ i)
                    If CurStep > AfAxialMax Then CurStep = AfAxialMax
                    If ((CurStep = 0) Or (CurStep > AfAxialMin)) And (CurStep - lastStep >= minStepSize) Then
                        rmStepList.Add "AFz", CurStep
                        lastStep = CurStep
                    End If
                Next
            Else
                StepMax = stepmaxAF
                If (StepMax > AfAxialMax) Or (StepMax > AfTransMax) Then
                    If AfTransMax > AfAxialMax Then StepMax = AfAxialMax Else StepMax = AfTransMax
                End If
                NumSteps = Log(StepMax) / Log(stepsizeAF)
                If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)
                lastStep = 0
                For i = 1 To NumSteps
                    CurStep = Int(stepsizeAF ^ i)
                    If (CurStep > AfAxialMax) Or (CurStep > AfTransMax) Then
                        If AfTransMax > AfAxialMax Then CurStep = AfAxialMax Else CurStep = AfTransMax
                    End If
                    If ((CurStep = 0) Or ((CurStep > AfAxialMin) And (CurStep > AfTransMin))) And (CurStep - lastStep >= minStepSize) Then
                        rmStepList.Add "AF", CurStep
                        lastStep = CurStep
                    End If
                Next i
            End If
        End If
    ' Rotational remanence magnetization acquisition & AF
    If chkRMAllWithRRM = vbChecked And EnableAF Then
        rmStepList.Add "AFmax", AfAxialMax
        StepSize = val(txtRMAllRRMrpsStep)
        StepMax = val(txtRMAllRRMrpsMax)
        field = val(txtRMAllRRMAFField)
        If StepSize = 0 Or StepSize > StepMax Then Exit Sub
        For i = 0 To StepMax Step StepSize
            rmStepList.Add "RRM", field, 0, i, 5
        Next i
        If chkRMAllRRMdoNegative = vbChecked Then
            rmStepList.Add "AFmax", AfAxialMax
            For i = 0 To -StepMax Step -StepSize
                rmStepList.Add "RRM", field, 0, i, 5
            Next i
        End If
        NumSteps = Log(stepmaxAF) / Log(stepsizeAF)
        If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)
        lastStep = 0
        For i = 1 To NumSteps
            CurStep = Int(stepsizeAF ^ i)
            If CurStep > AfAxialMax Then CurStep = AfAxialMax
            If ((CurStep = 0) Or (CurStep > AfAxialMin)) And (CurStep - lastStep >= minStepSize) Then
                rmStepList.Add "AFz", CurStep
                lastStep = CurStep
            End If
        Next i
    End If
    ' ARM acquisition & AF demag
    If chkRMAllARM And EnableARM And EnableAF Then
        rmStepList.Add "AFmax", AfAxialMax
        StepSize = val(txtRMAllARMStepSize)
        StepMax = val(txtRMAllARMStepMax)
        field = val(txtRMAllAFFieldForARM)
        If field > 0.75 * AfAxialMax Then field = 0.75 * AfAxialMax
        If StepSize = 0 Then Exit Sub
        On Error GoTo continuance
        If StepMax Mod StepSize <> 0 Then StepMax = StepMax - StepMax Mod StepSize
continuance:
        On Error GoTo 0
        For i = 0 To StepMax Step StepSize
            rmStepList.Add "ARM", field, i
        Next i
        NumSteps = (Log(stepmaxAF) / Log(stepsizeAF))
        If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)
        lastStep = 0
        For i = 1 To NumSteps
            CurStep = Int(stepsizeAF ^ i)
            If CurStep > AfAxialMax Then CurStep = AfAxialMax
            If ((CurStep = 0) Or (CurStep > AfAxialMin)) And (CurStep - lastStep >= minStepSize) Then
                rmStepList.Add "AFz", CurStep
                lastStep = CurStep
            End If
        Next i
     
        'Clean last of ARM field possibly left
        rmStepList.Add "AFmax", AfAxialMax
           
        ' IRM pulse and AF demag
        'Only do this sequence if the IRM's are enables
        If EnableAxialIRM Or EnableTransIRM Then
            
            rmStepList.Add "IRMz", 0
            rmStepList.Add "AFmax", AfAxialMax
            rmStepList.Add "IRMz", field
            lastStep = 0
            For i = 1 To NumSteps
                CurStep = Int(stepsizeAF ^ i)
                If CurStep > AfAxialMax Then CurStep = AfAxialMax
                If ((CurStep = 0) Or (CurStep > AfAxialMin)) And (CurStep - lastStep >= minStepSize) Then
                    rmStepList.Add "AFz", CurStep
                    lastStep = CurStep
                End If
            Next i
        End If
            
    End If
     ' IRM stepwise and final cleaning
    If chkRMAllIRM And (EnableAxialIRM Or EnableTransIRM) And EnableAF Then
        rmStepList.Add "AFmax", AfAxialMax
        rmStepList.Add "IRMz", 0, MeasureSusceptibility:=(chkTheWorksMeasSusc.value = Checked)
        rmStepList.Add "AFmax", AfAxialMax
        NumSteps = (Log(val(txtRMAllIRMMax)) / Log(stepsizeAF))
        If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)
        lastStep = 0
        For i = 1 To NumSteps
            CurStep = Int(stepsizeAF ^ i)
            If CurStep > PulseAxialMax Then
                
                    CurStep = PulseAxialMax
                    
            End If
            If ((CurStep = 0) Or (CurStep > PulseAxialMin)) And (CurStep - lastStep >= minStepSize) Then
                rmStepList.Add "IRMz", CurStep
                lastStep = CurStep
            End If
        Next i
        NumSteps = Log(stepmaxAF) / Log(stepsizeAF)
        If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)
        lastStep = 0
        For i = 1 To NumSteps
            CurStep = Int(stepsizeAF ^ i)
            If CurStep > AfAxialMax Then CurStep = AfAxialMax
            If ((CurStep = 0) Or (CurStep > AfAxialMin)) And (CurStep - lastStep >= minStepSize) Then
                rmStepList.Add "AFz", CurStep
                lastStep = CurStep
            End If
        Next i
        rmStepList.Add "AFmax", AfAxialMax
    End If
     ' IRM backfield DC demag
    If chkRMAllBackfieldDemag And EnableIRMBackfield And (EnableAxialIRM Or EnableTransIRM) Then
        NumSteps = (Log(val(txtRMAllIRMMax)) / Log(stepsizeAF))
        If NumSteps - Int(NumSteps) > 0.5 Then NumSteps = Int(NumSteps) + 1 Else NumSteps = Int(NumSteps)
        CurStep = Int(stepsizeAF ^ NumSteps)
        If CurStep > PulseAxialMax Then
                
            CurStep = PulseAxialMax
            
        End If
        If (CurStep = 0) Or (CurStep > PulseAxialMin) Then rmStepList.Add "IRMz", CurStep
       lastStep = CurStep
        For i = 1 To NumSteps - 1
            CurStep = -Int(stepsizeAF ^ i)
            If -CurStep > PulseAxialMax Then CurStep = -PulseAxialMax
            If ((CurStep = 0) Or (-CurStep > PulseAxialMin)) And (Abs(CurStep - lastStep) >= minStepSize) Then
                rmStepList.Add "IRMz", CurStep
                lastStep = CurStep
            End If
        Next i
        
        'Check to see if the AF module is enabled
        If EnableAF = True Then rmStepList.Add "AFmax", AfAxialMax
        
    End If
     refreshListDisplay
End Sub

Private Sub cmdSave_Click()
    Dim sfilen  As String           ' Path + filename of selected file
    ' Initialize the dialog box for Save
    dlgCommonDialog.FILTER = "Rockmag file (*.rmg)|*.rmg|All files (*.*)|*.*"
    dlgCommonDialog.DialogTitle = "Save to RMG File..."
    If FileExists(Prog_DefaultPath) Then
        dlgCommonDialog.InitDir = Prog_DefaultPath
    Else
        dlgCommonDialog.InitDir = "\"
    End If
    dlgCommonDialog.ShowSave
    sfilen = dlgCommonDialog.filename
    If LenB(sfilen) = 0 Then Exit Sub ' If we don't have a filename then don't processs it.
    SaveRMGRoutine (sfilen)
End Sub

'Sub EnableDisableControls
'
' Created: Jan. 2011
'  Author: I Hilburn
'
' Summary:  Enables or disables controls on the form based upon
'           the rockmag modules settings
'
'  Return:  None
'   Input:  None
'
' Effects: If EnableAF = False
'          Then all the controls using AF's will be unchecked and disabled
'          If EnableARM = False
'          Then all the controls using ARM will be unchecked and disabled
'          If EnableAxialIRM & EnableTransIRM both = False
'          Then all the IRM controls will be unchecked and disabled
'          If EnableSusceptibility = False
'          Then all the measure susceptibility check-boxes with be unchecked
'          and disabled
'          Otherwise, all the controls will be enabled and set to their default values
'          as stored in the INI file

Private Sub EnableDisableControls()
    
    'Need to uncheck all check boxes that correspond to disabled Rockmag modules
    
    'Susceptibility Module
    If EnableSusceptibility = False Then
        
        Me.chkTheWorksMeasSusc.value = Unchecked
        Me.chksusceptibility.value = Unchecked
        
    End If
    
    'ARM module (only)
    If EnableARM = False Then Me.chkRMAllARM.value = Unchecked
    
    'IRM module
    If EnableAxialIRM = False And _
       EnableTransIRM = False _
    Then
    
        Me.chkRMAllIRM.value = Unchecked
        Me.chkRMAllBackfieldDemag.value = Unchecked
        
    End If
    
    'IRM Backfield module
    If EnableIRMBackfield = False Then
    
        Me.chkRMAllBackfieldDemag.value = Unchecked
        
    End If
    
    'AF module
    If EnableAF = False Then
    
        'Uncheck all the ARM/IRM/RRM/AF controls - all depend on AF
        'Except the IRM DC backfield - this does not require AF's
        Me.chkRMAllNRM.value = Unchecked
        Me.chkRMAllRRMdoNegative.value = Unchecked
        Me.chkRMAllNRM3AxisAF.value = Unchecked
        Me.chkRMAllWithRRM.value = Unchecked
        Me.chkRMAllIRM.value = Unchecked
        Me.chkRMAllARM.value = Unchecked
        
    End If
    
    'Need to make set contols' enabled / disabled state +
    'checked or unchecked state based upon whether
    'the corresponding modules are enabled or disabled
       
    Me.chkTheWorksMeasSusc.Enabled = EnableSusceptibility
    Me.chkRMAllARM.Enabled = EnableARM And EnableAF
    Me.chkRMAllIRM.Enabled = (EnableAxialIRM Or EnableTransIRM) And EnableAF
    Me.chkRMAllBackfieldDemag.Enabled = EnableIRMBackfield And _
                                        (EnableAxialIRM Or EnableTransIRM)
    Me.chkRMAllNRM.Enabled = EnableAF
    Me.chkRMAllRRMdoNegative.Enabled = EnableAF
    Me.chkRMAllNRM3AxisAF.Enabled = EnableAF
    Me.chkRMAllWithRRM.Enabled = EnableAF
    Me.txtRMAllAFFieldForARM.Enabled = EnableARM And EnableAF
    Me.txtRMAllAFMax.Enabled = EnableAF
    Me.txtRMAllARMStepMax.Enabled = EnableARM And EnableAF
    Me.txtRMAllARMStepSize.Enabled = EnableARM And EnableAF
    Me.txtRMAllIRMMax.Enabled = (EnableAxialIRM Or EnableTransIRM)
    Me.txtRMAllLogFactor.Enabled = (EnableAxialIRM Or EnableTransIRM Or EnableAF)
            
    'Enable/Disable the Hawaiian Standard demag button
    Me.cmdHawaiianStd.Enabled = EnableAF
            
    'Always Enable the Rockmag the works button
    Me.cmdRockmagEverything.Enabled = True
    
    'Now Enable all of the ARM controls in the middle frame
    Me.txtARMBiasMax.Enabled = EnableARM And EnableAF
    Me.txtARMSteps.Enabled = EnableARM And EnableAF
    Me.txtAFfieldForARM.Enabled = EnableARM And EnableAF
    Me.txtBiasField.Enabled = EnableARM And EnableAF
    Me.txtSpinSpeed.Enabled = EnableAF
    Me.cmbARMStepSeqScale.Enabled = EnableARM And EnableAF
    Me.cmdAddARMStepSeq.Enabled = EnableARM And EnableAF
    
    'Now enable all the rest of the controls in the middle of the form
    Me.txtHoldTime.Enabled = EnableAF
    Me.txtLevel.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.txtStepMin.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.txtStepMax.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.txtStepSize.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.cmdAdd.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.cmdaddStepSeq.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.chkMeasure.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.chksusceptibility.Enabled = EnableSusceptibility
    Me.cmbStepType.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    Me.cmbStepSeq.Enabled = EnableAF Or EnableAxialIRM Or EnableTransIRM
    
End Sub

Private Sub Form_Activate()

     'Disable all the controls on the form used to set AF steps
    EnableDisableControls
                
    'Change the combo Step controls to allow only non-AF steps
    LoadStepTypeSeqComboBoxes
    
    'Reset the max AF & IRM values
    Me.txtRMAllIRMMax = modConfig.PulseAxialMax
    Me.txtRMAllAFMax = modConfig.AfAxialMax
        
End Sub

Private Sub Form_Load()
    
    Dim colX As ColumnHeader ' Declare variable.
    
    '(March 11, 2011 - I Hilburn)
    'Update all the form's icons with the program Ico file
    SetFormIcon Me
    
    SequenceReady = False
    Set rmStepList = New RockmagSteps
    Set colX = lvwSteps.ColumnHeaders.Add(1)
    colX.text = "#"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSteps.ColumnHeaders.Add(2)
    colX.text = "Step Type"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSteps.ColumnHeaders.Add(3)
    colX.text = "Level (G)"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSteps.ColumnHeaders.Add(4)
    colX.text = "Bias Field (G)"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSteps.ColumnHeaders.Add(5)
    colX.text = "Spin (RPS)"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSteps.ColumnHeaders.Add(6)
    colX.text = "Hold (sec)"
    colX.Width = Me.TextWidth(colX.text & " ")
    Set colX = lvwSteps.ColumnHeaders.Add(7)
    colX.text = "Measure?"
    colX.Width = Me.TextWidth(colX.text)
    Set colX = lvwSteps.ColumnHeaders.Add(8)
    colX.text = "Suscep?"
    colX.Width = Me.TextWidth(colX.text)
    Set colX = lvwSteps.ColumnHeaders.Add(9) ' (November 2007 L Carporzen) Remarks column in RMG
    colX.text = "Remarks"
    colX.Width = Me.TextWidth(colX.text)

    LoadStepTypeSeqComboBoxes

    cmbStepSeqScale.Clear
    cmbStepSeqScale.AddItem "Linear"
    cmbStepSeqScale.AddItem "Log"
    cmbStepSeqScale.ListIndex = 0
    cmbARMStepSeqScale.Clear
    cmbARMStepSeqScale.AddItem "Linear"
    cmbARMStepSeqScale.AddItem "Log"
    cmbARMStepSeqScale.ListIndex = 0
    chkMeasure.value = Checked
    loadDefaults
    
    'Enable or disable rock-mag step controls based upon
    'whether or not the corresponding modules are enabled
    EnableDisableControls
        
    'Reset the max AF & IRM values
    Me.txtRMAllIRMMax = modConfig.PulseAxialMax
    Me.txtRMAllAFMax = modConfig.AfAxialMax
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Me.Height = 9060
        Me.Width = 9690
    End If
End Sub

Private Sub ImportRMGRoutine(ByVal filen As String)
    Dim filenum As Integer
    Dim whole_file As String
    Dim lines As Variant
    Dim the_array As Variant
    Dim num_rows As Long
    Dim r As Long
    Dim readStepType As String
    Dim readStepLevel As Double
    Dim readBias As Double
    Dim readSpin As Double
    Dim readHold As Double
    Dim readMeas As Double
    Dim readSusceptibility As Double
    filenum = FreeFile
    Open filen For Input As #filenum
    whole_file = Input$(LOF(filenum), #filenum)
    Close #filenum
    lines = Split(whole_file, vbCrLf)
    num_rows = UBound(lines)
    ReDim the_array(num_rows)
    For r = 0 To num_rows
        the_array(r) = Split(lines(r), ",")
    Next r
    For r = 0 To UBound(the_array) - 1
        If UBound(the_array(r)) > 9 Then
            If Left$(the_array(r)(0), 1) = " " Or Left$(the_array(r)(0), 5) = "Level" Or Left$(the_array(r)(0), 12) = "Instrument: " Or Left$(the_array(r)(0), 6) = "Time: " Then
            Else
            'If val(the_array(r)(5)) <> 0 Then
                readStepType = the_array(r)(0)
                If the_array(r)(0) = "UAFZ" Then readStepType = "UAFZ1" 'correct the old UAFZ
                readStepLevel = val(the_array(r)(1))
                readBias = val(the_array(r)(2))
                readSpin = val(the_array(r)(3))
                readHold = val(the_array(r)(4))
                readMeas = val(the_array(r)(5))
                readSusceptibility = val(the_array(r)(8))
                rmStepList.Add readStepType, readStepLevel, readBias, readSpin, readHold, (readMeas <> 0), (readSusceptibility <> 0)
            End If
        End If
    Next r
    refreshListDisplay
End Sub

Private Sub loadDefaults()
    chkRMAllNRM.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkRMAllNRM", "0", Prog_INIFile))
    chkRMAllNRM3AxisAF.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkRMAllNRM3AxisAF", "0", Prog_INIFile))
    chkRMAllWithRRM.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkRMAllWithRRM", "0", Prog_INIFile))
    chkRMAllRRMdoNegative.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkRMAllRRMdoNegative", "0", Prog_INIFile))
    chkRMAllARM.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkRMAllARM", "0", Prog_INIFile))
    chkRMAllIRM.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkRMAllIRM", "0", Prog_INIFile))
    chkRMAllBackfieldDemag.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkRMAllBackfieldDemag", "0", Prog_INIFile))
    chkMeasure.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkMeasure", "1", Prog_INIFile))
    chkTheWorksMeasSusc.value = Unchecked
    chksusceptibility.value = val(Config_GetFromINI("RockmagRoutineDefaults", "chkSusceptibility", "0", Prog_INIFile))
    txtRMAllRRMrpsStep.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllRRMrpsStep", vbNullString, Prog_INIFile)
    txtRMAllRRMrpsMax.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllRRMrpsMax", vbNullString, Prog_INIFile)
    txtRMAllRRMAFField.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllRRMAFField", vbNullString, Prog_INIFile)
    txtRMAllARMStepSize.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllARMStepSize", vbNullString, Prog_INIFile)
    txtRMAllARMStepMax.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllARMStepMax", vbNullString, Prog_INIFile)
    txtRMAllAFFieldForARM.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllAFFieldForARM", vbNullString, Prog_INIFile)
    txtRMAllLogFactor.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllLogFactor", vbNullString, Prog_INIFile)
    txtRMAllMinStepSize.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllMinStepSize", vbNullString, Prog_INIFile)
    txtRMAllAFMax.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllAFMax", vbNullString, Prog_INIFile)
    txtRMAllIRMMax.text = Config_GetFromINI("RockmagRoutineDefaults", "txtRMAllIRMMax", vbNullString, Prog_INIFile)
    txtStepSize.text = Config_GetFromINI("RockmagRoutineDefaults", "txtStepSize", vbNullString, Prog_INIFile)
    txtStepMax.text = Config_GetFromINI("RockmagRoutineDefaults", "txtStepMax", vbNullString, Prog_INIFile)
    txtARMSteps.text = Config_GetFromINI("RockmagRoutineDefaults", "txtARMSteps", vbNullString, Prog_INIFile)
    txtARMBiasMax.text = Config_GetFromINI("RockmagRoutineDefaults", "txtARMBiasMax", vbNullString, Prog_INIFile)
    txtAFfieldForARM.text = Config_GetFromINI("RockmagRoutineDefaults", "txtAFfieldForARM", vbNullString, Prog_INIFile)
    cmbStepSeq.text = Config_GetFromINI("RockmagRoutineDefaults", "cmbStepSeq", "AF", Prog_INIFile)
    cmbStepSeqScale.text = Config_GetFromINI("RockmagRoutineDefaults", "cmbStepSeqScale", "Log", Prog_INIFile)
    cmbARMStepSeqScale.text = Config_GetFromINI("RockmagRoutineDefaults", "cmbARMStepSeqScale", "Linear", Prog_INIFile)
    cmbStepType.text = Config_GetFromINI("RockmagRoutineDefaults", "cmbStepType", "AF", Prog_INIFile)
End Sub

Private Sub LoadStepTypeSeqComboBoxes()

    cmbStepType.Enabled = True
    cmbStepSeq.Enabled = True

    If EnableAF = True And _
       EnableARM = True And _
       (EnableAxialIRM = True Or _
        EnableTransIRM = True) _
    Then
    
        cmbStepType.Clear
        cmbStepType.AddItem "AF"
        cmbStepType.AddItem "AFz"
        cmbStepType.AddItem "AFmax"
        cmbStepType.AddItem "UAF"
        cmbStepType.AddItem "UAFZ1"
        cmbStepType.AddItem "UAFX1"
        cmbStepType.AddItem "UAFX2"
        cmbStepType.AddItem "aTAFX" ' (February 2010 L Carporzen) Measure the TAF and uncorrect them in sample file
        cmbStepType.AddItem "aTAFY"
        cmbStepType.AddItem "aTAFZ"
        cmbStepType.AddItem "GRM AF"
        cmbStepType.AddItem "ARM"
        cmbStepType.AddItem "IRMz"
        cmbStepType.AddItem "NRM"
        cmbStepType.AddItem "VRM"
        cmbStepType.AddItem "RRM"
        cmbStepType.AddItem "RRMz"
        cmbStepType.AddItem "X"
        cmbStepSeq.Clear
        cmbStepSeq.AddItem "AF"
        cmbStepSeq.AddItem "AFz"
        cmbStepSeq.AddItem "IRMz"
        cmbStepSeq.ListIndex = 0
        
    ElseIf EnableAF = True And _
           EnableARM = False And _
           (EnableAxialIRM = True Or _
            EnableTransIRM = True) _
    Then
    
        cmbStepType.Clear
        cmbStepType.AddItem "AF"
        cmbStepType.AddItem "AFz"
        cmbStepType.AddItem "AFmax"
        cmbStepType.AddItem "UAF"
        cmbStepType.AddItem "GRM AF"
        cmbStepType.AddItem "IRMz"
        cmbStepType.AddItem "NRM"
        cmbStepType.AddItem "RRM"
        cmbStepType.AddItem "RRMz"
        cmbStepType.AddItem "X"
        cmbStepSeq.Clear
        cmbStepSeq.AddItem "AF"
        cmbStepSeq.AddItem "AFz"
        cmbStepSeq.AddItem "IRMz"
        cmbStepSeq.ListIndex = 0
        
    ElseIf EnableAF = True And _
           EnableARM = True And _
           (EnableAxialIRM = False And _
            EnableTransIRM = False) _
    Then

        cmbStepType.Clear
        cmbStepType.AddItem "AF"
        cmbStepType.AddItem "AFz"
        cmbStepType.AddItem "AFmax"
        cmbStepType.AddItem "UAF"
        cmbStepType.AddItem "GRM AF"
        cmbStepType.AddItem "ARM"
        cmbStepType.AddItem "NRM"
        cmbStepType.AddItem "RRM"
        cmbStepType.AddItem "RRMz"
        cmbStepType.AddItem "X"
        cmbStepSeq.Clear
        cmbStepSeq.AddItem "AF"
        cmbStepSeq.AddItem "AFz"
        cmbStepSeq.ListIndex = 0
        
    ElseIf EnableAF = True And _
           EnableARM = False And _
           (EnableAxialIRM = False And _
            EnableTransIRM = False) _
    Then

        cmbStepType.Clear
        cmbStepType.AddItem "AF"
        cmbStepType.AddItem "AFz"
        cmbStepType.AddItem "AFmax"
        cmbStepType.AddItem "UAF"
        cmbStepType.AddItem "GRM AF"
        cmbStepType.AddItem "NRM"
        cmbStepType.AddItem "RRM"
        cmbStepType.AddItem "RRMz"
        cmbStepType.AddItem "X"
        cmbStepSeq.Clear
        cmbStepSeq.AddItem "AF"
        cmbStepSeq.AddItem "AFz"
        cmbStepSeq.ListIndex = 0
    
    ElseIf EnableAF = False And _
           (EnableAxialIRM = True Or _
            EnableTransIRM = True) _
    Then
    
        cmbStepType.Clear
        cmbStepType.AddItem "IRMz"
        cmbStepType.AddItem "NRM"
        
        cmbStepSeq.Clear
        cmbStepSeq.AddItem "IRMz"
        cmbStepSeq.ListIndex = 0

    ElseIf EnableAF = False And _
           (EnableAxialIRM = False And _
            EnableTransIRM = False) _
    Then
    
        cmbStepType.Enabled = False
        cmbStepSeq.Enabled = False
    
    End If

End Sub

Private Sub lvwSteps_MouseDown(Button As Integer, _
      Shift As Integer, X As Single, Y As Single)
' (September 2007 L Carporzen) RockMag lines editable by a right click
   If rmStepList.Count = 0 Then Exit Sub
   If Button = vbRightButton And Not lvwSteps.SelectedItem.index = 0 Then
    rmStepNumber = lvwSteps.SelectedItem.index
    With rmStepList.Item(rmStepNumber)
        cmbStepType = .StepType
        txtLevel = .Level
        txtBiasField = .BiasField
        txtSpinSpeed = .SpinSpeed
        txtHoldTime = .HoldTime
        If .Measure Then chkMeasure.value = Checked Else chkMeasure.value = Unchecked
        If .MeasureSusceptibility Then chksusceptibility.value = Checked Else chksusceptibility.value = Unchecked
        txtRemarks = .Remarks ' (November 2007 L Carporzen) Remarks column in RMG
    End With
    cmdReplace.Visible = True ' Replace the button "Add" by a "Replace" one
    cmdAdd.Visible = False ' Replace the button "Add" by a "Replace" one
    MsgBox "You can edit the parameters of the line " & rmStepNumber & ". Click on replace when done."
   End If
End Sub

Public Sub refreshListDisplay()
    Dim i As Integer
    Dim curItem As ListItem
    lvwSteps.ListItems.Clear
    With rmStepList
        If .Count = 0 Then Exit Sub
        For i = 1 To .Count
            With .Item(i)
                Set curItem = lvwSteps.ListItems.Add(i, .key)
                curItem.text = i
                curItem.SubItems(1) = .StepType
                curItem.SubItems(2) = .Level
                curItem.SubItems(3) = .BiasField
                curItem.SubItems(4) = .SpinSpeed
                curItem.SubItems(5) = .HoldTime
                If .Measure Then curItem.SubItems(6) = "Y" Else curItem.SubItems(6) = "N"
                If .MeasureSusceptibility Then curItem.SubItems(7) = "Y" Else curItem.SubItems(7) = "N"
                curItem.SubItems(8) = .Remarks ' (November 2007 L Carporzen) Remarks column in RMG
            End With
        Next i
    End With
End Sub

Private Sub saveDefaults()
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllNRM", str(chkRMAllNRM.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllNRM3AxisAF", str(chkRMAllNRM3AxisAF.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllWithRRM", str(chkRMAllWithRRM.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllRRMdoNegative", str(chkRMAllRRMdoNegative.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllARM", str(chkRMAllARM.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllIRM", str(chkRMAllIRM.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllBackfieldDemag", str(chkRMAllBackfieldDemag.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkMeasure", str(chkMeasure.value)
    Config_SaveSetting "RockmagRoutineDefaults", "chkRMAllRRMdoNegative", str(chkRMAllRRMdoNegative.value)
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllRRMrpsStep", txtRMAllRRMrpsStep.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllRRMrpsMax", txtRMAllRRMrpsMax.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllRRMAFField", txtRMAllRRMAFField.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllARMStepSize", txtRMAllARMStepSize.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllARMStepMax", txtRMAllARMStepMax.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllAFFieldForARM", txtRMAllAFFieldForARM.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllLogFactor", txtRMAllLogFactor.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllMinStepSize", txtRMAllMinStepSize.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllAFMax", txtRMAllAFMax.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtRMAllIRMMax", txtRMAllIRMMax.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtStepSize", txtStepSize.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtStepMax", txtStepMax.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtARMSteps", txtARMSteps.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtARMBiasMax", txtARMBiasMax.text
    Config_SaveSetting "RockmagRoutineDefaults", "txtAFfieldForARM", txtAFfieldForARM.text
    Config_SaveSetting "RockmagRoutineDefaults", "cmbStepSeq", cmbStepSeq.text
    Config_SaveSetting "RockmagRoutineDefaults", "cmbStepSeqScale", cmbStepSeqScale.text
    Config_SaveSetting "RockmagRoutineDefaults", "cmbARMStepSeqScale", cmbARMStepSeqScale.text
    Config_SaveSetting "RockmagRoutineDefaults", "cmbStepType", cmbStepType.text
End Sub

Private Sub SaveRMGRoutine(ByVal filen As String)
    Dim FilePath As String
    Dim filenum As Integer
    Dim i As Integer
    Dim j As Double
    Dim curItem As ListItem
    FilePath = filen
    filenum = FreeFile
    If FileExists(FilePath) = True Then
        If MsgBox("Would you like to replace the existing file?", vbYesNo) = vbNo Then Exit Sub
    End If
    ' Create the file new even if it exist
    Open FilePath For Output As #filenum
    Print #filenum, "Level"; ",";
    Print #filenum, "Bias Field (G)"; ",";
    Print #filenum, "Spin Speed (rps)"; ",";
    Print #filenum, "Hold Time (s)"; ",";
    Print #filenum, "Mz (emu)"; ",";
    Print #filenum, "Std. Dev. Z"; ",";
    Print #filenum, "Mz/Vol"; ",";
    Print #filenum, "Moment Susceptibility (emu/Oe)"; ",";
    Print #filenum, "Mx (emu)"; ",";
    Print #filenum, "Std. Dev. X"; ",";
    Print #filenum, "My (emu)"; ",";
    Print #filenum, "Std. Dev. Y"; " "
    Close #filenum
    Open FilePath For Append As #filenum
    With rmStepList
        If .Count = 0 Then Exit Sub
        For i = 1 To .Count
            With .Item(i)
                Print #filenum, .StepType; ",";
                Print #filenum, .Level; ",";
                Print #filenum, .BiasField; ",";
                Print #filenum, .SpinSpeed; ",";
                Print #filenum, .HoldTime; ",";
                j = 1.00000000000001E-09
                If .Measure Then Print #filenum, j; ","; Else Print #filenum, " 0 "; ","; ' Measure
                If .Measure Then Print #filenum, j; ","; Else Print #filenum, " 0 "; ",";
                If .Measure Then Print #filenum, j; ","; Else Print #filenum, " 0 "; ",";
                j = 1.00000000001
                If .MeasureSusceptibility Then Print #filenum, j; ","; Else Print #filenum, " 0 "; ","; ' Susceptibility
                j = 1.00000000000001E-09
                If .Measure Then Print #filenum, j; ","; Else Print #filenum, " 0 "; ",";
                If .Measure Then Print #filenum, j; ","; Else Print #filenum, " 0 "; ",";
                If .Measure Then Print #filenum, j; ","; Else Print #filenum, " 0 "; ",";
                If .Measure Then Print #filenum, j; Else Print #filenum, " 0 ";
                Print #filenum, " "
            End With
        Next i
    End With
    Close #filenum
End Sub

