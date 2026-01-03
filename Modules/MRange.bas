Attribute VB_Name = "MRange"
Option Explicit

Private Type VBGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data5(0 To 7) As Byte
End Type

Private Type TIRange
    pVTable As LongPtr ' First element in an object always is a pointer to it's VTable
    refCnt  As Long    ' the reference counter
    Minimum As Double
    Maximum As Double
End Type

Private Type TRangeVTable
    '0 to 2 IUnknown
    '3 to 6 IDispatch
    '7 to 9 IRange.IsIn, IRange.IsOut, IRange.ToStr
    Funcs(0 To 9) As LongPtr
End Type

Private m_IRangeExMinExMaxVTable  As TRangeVTable
Private m_pIRangeExMinExMaxVTable As LongPtr

Private m_IRangeExMinInMaxVTable  As TRangeVTable
Private m_pIRangeExMinInMaxVTable As LongPtr

Private m_IRangeInMinExMaxVTable  As TRangeVTable
Private m_pIRangeInMinExMaxVTable As LongPtr

Private m_IRangeInMinInMaxVTable  As TRangeVTable
Private m_pIRangeInMinInMaxVTable As LongPtr

Public Sub InitIRangeVTable()
    Dim i As Long
    With m_IRangeExMinExMaxVTable
        .Funcs(i) = MPtr.FncPtr(AddressOf IUnknown_FncQueryInterface):  i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IUnknown_SubAddRef):          i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IUnknown_SubRelease):         i = i + 1
        
        .Funcs(i) = MPtr.FncPtr(AddressOf IDispatch_get_TypeInfoCount): i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IDispatch_get_TypeInfo):      i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IDispatch_get_IDsOfNames):    i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IDispatch_FncInvoke):         i = i + 1
    End With
    m_IRangeExMinInMaxVTable = m_IRangeExMinExMaxVTable
    m_IRangeInMinExMaxVTable = m_IRangeExMinInMaxVTable
    m_IRangeInMinInMaxVTable = m_IRangeInMinExMaxVTable
    With m_IRangeExMinExMaxVTable
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeExMinExMax_IsIn):       i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeExMinExMax_IsOut):      i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeExMinExMax_ToStr):      i = i + 1
    End With
    i = 7
    With m_IRangeExMinInMaxVTable
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeExMinInMax_IsIn):       i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeExMinInMax_IsOut):      i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeExMinInMax_ToStr):      i = i + 1
    End With
    i = 7
    With m_IRangeInMinExMaxVTable
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeInMinInMax_IsIn):       i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeInMinInMax_IsOut):      i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeInMinInMax_ToStr):      i = i + 1
    End With
    i = 7
    With m_IRangeInMinInMaxVTable
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeInMinInMax_IsIn):       i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeInMinInMax_IsOut):      i = i + 1
        .Funcs(i) = MPtr.FncPtr(AddressOf IRangeInMinInMax_ToStr):      i = i + 1
    End With
        '.Funcs(i) = MPtr.FncPtr(AddressOf IRange_FncIsIn):              i = i + 1
        '.Funcs(i) = MPtr.FncPtr(AddressOf IRange_FncIsOut):             i = i + 1
        '.Funcs(i) = MPtr.FncPtr(AddressOf IRange_FncToStr):             i = i + 1
End Sub

Public Function IRangeExMinExMax(ByVal RangeMin As Double, ByVal RangeMax As Double) As IRange
    
End Function

Public Function IRangeExMinInMax(ByVal RangeMin As Double, ByVal RangeMax As Double) As IRange
    
End Function

Public Function IRangeInMinExMax(ByVal RangeMin As Double, ByVal RangeMax As Double) As IRange
    
End Function

Public Function IRangeInMinInMax(ByVal RangeMin As Double, ByVal RangeMax As Double) As IRange
    
End Function

' v ############################## v '    IUnkown    ' v ############################## v '
Private Function IUnknown_FncQueryInterface(this As TIRange, riid As VBGUID, pvObj As LongPtr) As LongPtr
    '
End Function

Private Function IUnknown_SubAddRef(this As TIRange) As Long
    'Debug.Print "IUnknown_SubAddRef"
    ' now we add one reference
    With this
        .refCnt = .refCnt + 1
    End With
End Function

Private Function IUnknown_SubRelease(this As TIRange) As Long
    'Debug.Print "IUnknown_SubRelease"
    ' now we subtract one reference
    With this
        .refCnt = .refCnt - 1
    End With
    'If this.refCnt = 0 Then 'cleanup
End Function
' ^ ############################## ^ '    IUnkown    ' ^ ############################## ^ '

' v ############################## v '   IDispatch   ' v ############################## v '
Private Function IDispatch_get_TypeInfoCount(this As TIRange) As Long
    'Debug.Print "IDispatch_get_TypeInfoCount"
End Function
Private Function IDispatch_get_TypeInfo(this As TIRange) As Long
    'Debug.Print "IDispatch_get_TypeInfo"
End Function
Private Function IDispatch_get_IDsOfNames(this As TIRange) As Long
    'Debug.Print "IDispatch_get_IDsOfNames"
End Function
Private Function IDispatch_FncInvoke(this As TIRange) As Long
    'Debug.Print "IDispatch_FncInvoke"
End Function
' ^ ############################## ^ '   IDispatch   ' ^ ############################## ^ '

' v ############################## v '    IRange     ' v ############################## v '
' v ############################## v '  ExMinExMax   ' v ############################## v '
Private Function IRangeExMinExMax_IsIn(this As TIRange, ByVal Value As Double) As Boolean
    'returns true if Value is inside the range otherwise false
    With this
        IRangeExMinExMax_IsIn = .Minimum < Value And Value < .Maximum
    End With
End Function
Private Function IRangeExMinExMax_IsOut(this As TIRange, ByVal Value As Double) As Integer
    'returns 0 / false if Value is "in" resp is "not out" the range
    'returns -1 / true if Value falls below min
    'returns 1 if Value exceeds max
    With this
        IRangeExMinExMax_IsOut = IIf(Value <= .Minimum Or .Maximum <= Value, IIf(Value <= .Minimum, -1, 1), 0)
    End With
End Function
Private Function IRangeExMinExMax_ToStr(this As TIRange, Optional ByVal FormatAsTyp As VbVarType = vbDouble) As String
    With this
        IRangeExMinExMax_ToStr = "] " & FormatRange(FormatAsTyp, .Minimum, .Maximum) & " [" & ": " & ERangeType_ToStr(ERangeType.ExMinExMax)
    End With
End Function
' ^ ############################## ^ '  ExMinExMax   ' ^ ############################## ^ '
' v ############################## v '  ExMinInMax   ' v ############################## v '
Private Function IRangeExMinInMax_IsIn(this As TIRange, ByVal Value As Double) As Boolean
    'returns true if Value is inside the range otherwise false
    With this
        IRangeExMinInMax_IsIn = .Minimum < Value And Value <= .Maximum
    End With
End Function
Private Function IRangeExMinInMax_IsOut(this As TIRange, ByVal Value As Double) As Integer
    'returns 0 / false if Value is "in" resp is "not out" the range
    'returns -1 / true if Value falls below min
    'returns 1 if Value exceeds max
    With this
        IRangeExMinInMax_IsOut = IIf(Value <= .Minimum Or .Maximum < Value, IIf(Value <= .Minimum, -1, 1), 0)
    End With
End Function
Private Function IRangeExMinInMax_ToStr(this As TIRange, Optional ByVal FormatAsTyp As VbVarType = vbDouble) As String
    With this
        IRangeExMinInMax_ToStr = "] " & FormatRange(FormatAsTyp, .Minimum, .Maximum) & " ]" & ": " & ERangeType_ToStr(ERangeType.ExMinInMax)
    End With
End Function
' ^ ############################## ^ '  ExMinInMax   ' ^ ############################## ^ '
' v ############################## v '  InMinExMax   ' v ############################## v '
Private Function IRangeInMinExMax_IsIn(this As TIRange, ByVal Value As Double) As Boolean
    'returns true if Value is inside the range otherwise false
    With this
        IRangeInMinExMax_IsIn = .Minimum <= Value And Value < .Maximum
    End With
End Function
Private Function IRangeInMinExMax_IsOut(this As TIRange, ByVal Value As Double) As Integer
    'returns 0 / false if Value is "in" resp is "not out" the range
    'returns -1 / true if Value falls below min
    'returns 1 if Value exceeds max
    With this
        IRangeInMinExMax_IsOut = IIf(Value < .Minimum Or .Maximum <= Value, IIf(Value < .Minimum, -1, 1), 0)
    End With
End Function
Private Function IRangeInMinExMax_ToStr(this As TIRange, Optional ByVal FormatAsTyp As VbVarType = vbDouble) As String
    With this
        IRangeInMinExMax_ToStr = "[ " & FormatRange(FormatAsTyp, .Minimum, .Maximum) & " [" & ": " & ERangeType_ToStr(ERangeType.InMinExMax)
    End With
End Function
' ^ ############################## ^ '  InMinExMax   ' ^ ############################## ^ '
' v ############################## v '  InMinInMax   ' v ############################## v '
Private Function IRangeInMinInMax_IsIn(this As TIRange, ByVal Value As Double) As Boolean
    'returns true if Value is inside the range otherwise false
    With this
        IRangeInMinInMax_IsIn = .Minimum <= Value And Value <= .Maximum
    End With
End Function
Private Function IRangeInMinInMax_IsOut(this As TIRange, ByVal Value As Double) As Integer
    'returns 0 / false if Value is "in" resp is "not out" the range
    'returns -1 / true if Value falls below min
    'returns 1 if Value exceeds max
    With this
        IRangeInMinInMax_IsOut = IIf(Value < .Minimum Or .Maximum < Value, IIf(Value < .Minimum, -1, 1), 0)
    End With
End Function
Private Function IRangeInMinInMax_ToStr(this As TIRange, Optional ByVal FormatAsTyp As VbVarType = vbDouble) As String
    With this
        IRangeInMinInMax_ToStr = "[ " & FormatRange(FormatAsTyp, .Minimum, .Maximum) & " ]" & ": " & ERangeType_ToStr(ERangeType.InMinInMax)
    End With
End Function
' ^ ############################## ^ '  InMinInMax   ' ^ ############################## ^ '
' ^ ############################## ^ '    IRange     ' ^ ############################## ^ '
