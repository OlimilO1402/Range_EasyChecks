VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangeInMinInMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRange
Private m_Min As Double
Private m_Max As Double

Friend Sub New_(MinValue As Double, MaxValue As Double)
    m_Min = MinValue: m_Max = MaxValue
    If m_Min <= m_Max Then Exit Sub
    m_Min = MaxValue: m_Max = MinValue
End Sub

'Copy to MNew:
'Public Function RangeInMinInMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeInMinInMax
'    Set RangeInMinInMax = New RangeInMinInMax: RangeInMinInMax.New_ MinValue, MaxValue
'End Function

Private Function IRange_IsIn(ByVal Value As Double) As Boolean
    'returns true if Value is inside the range otherwise false
    IRange_IsIn = m_Min <= Value And Value <= m_Max
End Function

Private Function IRange_IsOut(ByVal Value As Double) As Integer
    'we also want to return if the Value falls below Min or if it exceeds Max
    'returns 0 / false if Value is "in" resp is "not out" the range
    'returns -1 / true if Value falls below min
    'returns 1 if Value exceeds max
    IRange_IsOut = IIf(Value < m_Min Or m_Max < Value, IIf(Value < m_Min, -1, 1), 0)
End Function

Private Function IRange_ToStr(Optional ByVal FormatAsTyp As VbVarType = vbDouble) As String
    IRange_ToStr = "[ " & FormatRange(FormatAsTyp, m_Min, m_Max) & " ]" & ": " & ERangeType_ToStr(ERangeType.InMinInMax)
End Function
