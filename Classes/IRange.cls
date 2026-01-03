VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function IsIn(ByVal Value As Double) As Boolean:  End Function
Public Function IsOut(ByVal Value As Double) As Integer: End Function
Public Function ToStr(Optional ByVal FormatAsTyp As VbVarType = vbDouble) As String: End Function
