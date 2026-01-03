VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRange
Private m_Min As Double
Private m_Max As Double
Private m_Typ As ERangeType

Friend Sub New_(MinValue As Double, MaxValue As Double, RangeType As ERangeType)
    m_Min = MinValue: m_Max = MaxValue: m_Typ = RangeType
End Sub

'Public Sub SetRange(ByVal MinValue As Double, ByVal MaxValue As Double)
'    m_MinValue = MinValue: m_MaxValue = MaxValue
'End Sub
'Public Property Get RangeType() As ERangeType
'    RangeType = m_RangeTyp
'End Property
'Public Property Let RangeType(ByVal Value As ERangeType)
'    m_RangeTyp = Value
'End Property

Public Function IsIn(ByVal Value As Double) As Boolean
    'returns true if Value is inside the range otherwise false
    Select Case m_Typ
    Case ERangeType.ExMinExMax: IsIn = m_Min < Value And Value < m_Max
    Case ERangeType.ExMinInMax: IsIn = m_Min < Value And Value <= m_Max
    Case ERangeType.InMinExMax: IsIn = m_Min <= Value And Value < m_Max
    Case ERangeType.InMinInMax: IsIn = m_Min <= Value And Value <= m_Max
    End Select
End Function

Public Function IsOut(ByVal Value As Double) As Integer
    'we also want to return if the Value falls below Min or if it exceeds Max
    'returns 0 / false if Value is "in" resp is "not out" the range
    'returns -1 / true if Value falls below min
    'returns 1 if Value exceeds max
    Select Case m_Typ
    Case ERangeType.ExMinExMax: 'IsOut = Value <= m_Min Or m_Max <= Value
                                IsOut = IIf(Value <= m_Min Or m_Max <= Value, IIf(Value <= m_Min, -1, 1), 0)
    Case ERangeType.ExMinInMax: 'IsOut = Value <= m_Min Or m_Max < Value
                                IsOut = IIf(Value <= m_Min Or m_Max < Value, IIf(Value <= m_Min, -1, 1), 0)
    Case ERangeType.InMinExMax: 'IsOut = Value < m_Min Or m_Max <= Value
                                IsOut = IIf(Value < m_Min Or m_Max <= Value, IIf(Value < m_Min, -1, 1), 0)
    Case ERangeType.InMinInMax: 'IsOut = Value < m_Min Or m_Max < Value
                                IsOut = IIf(Value < m_Min Or m_Max < Value, IIf(Value < m_Min, -1, 1), 0)
    End Select
End Function

'Public Enum ERangeType
'    ExMinExMax = 0 ' 0 0
'    ExMinInMax = 1 ' 0 1
'    InMinExMax = 2 ' 1 0
'    InMinInMax = 3 ' 1 1
'End Enum

Public Function ToStr(Optional FormatAsTyp As VbVarType = vbDouble) As String
    Dim pl As String: pl = IIf((m_Typ And 2) = 2, "[ ", "] ")
    Dim pr As String: pr = IIf((m_Typ And 1) = 1, " ]", " [")
    ToStr = pl & FormatRange(FormatAsTyp, m_Min, m_Max) & pr & ": " & ERangeType_ToStr(m_Typ)
End Function

Private Function IRange_IsIn(ByVal Value As Double) As Boolean
    IRange_IsIn = Me.IsIn(Value)
End Function

Private Function IRange_IsOut(ByVal Value As Double) As Integer
    IRange_IsOut = Me.IsOut(Value)
End Function

Private Function IRange_ToStr(Optional ByVal FormatAsTyp As VbVarType = VbVarType.vbDouble) As String
    IRange_ToStr = Me.ToStr(FormatAsTyp)
End Function
