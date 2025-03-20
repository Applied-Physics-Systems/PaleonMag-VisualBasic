VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADwin_DOChan_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarByte_ID As Byte 'local copy
Private mvarIs_High As Tristate 'local copy
Public Property Let Is_High(ByVal vData As Tristate)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Is_High = 5
    mvarIs_High = vData
End Property


Public Property Get Is_High() As Tristate
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Is_High
    Is_High = mvarIs_High
End Property



Public Property Let Byte_ID(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Byte_ID = 5
    mvarByte_ID = vData
End Property


Public Property Get Byte_ID() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Byte_ID
    Byte_ID = mvarByte_ID
End Property



