Attribute VB_Name = "MNew"
Option Explicit

Public Enum ERangeType
    ExMinExMax = 0 ' 0 0
    ExMinInMax = 1 ' 0 1
    InMinExMax = 2 ' 1 0
    InMinInMax = 3 ' 1 1
End Enum

Public Function ERangeType_ToStr(e As ERangeType) As String
    Dim s As String
    Select Case e
    Case ERangeType.ExMinExMax: s = "Minimum excluded and maximum excluded"
    Case ERangeType.ExMinInMax: s = "Minimum excluded but maximum included"
    Case ERangeType.InMinExMax: s = "Minimum included but maximum excluded"
    Case ERangeType.InMinInMax: s = "Minimum included and maximum included"
    End Select
    ERangeType_ToStr = s
End Function

Public Function Range(ByVal MinValue As Double, ByVal MaxValue As Double, ByVal RangeType As ERangeType) As Range
    Set Range = New Range: Range.New_ MinValue, MaxValue, RangeType
End Function

Public Function RangeExMinExMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeExMinExMax
    Set RangeExMinExMax = New RangeExMinExMax: RangeExMinExMax.New_ MinValue, MaxValue
End Function
Public Function RangeExMinInMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeExMinInMax
    Set RangeExMinInMax = New RangeExMinInMax: RangeExMinInMax.New_ MinValue, MaxValue
End Function
Public Function RangeInMinExMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeInMinExMax
    Set RangeInMinExMax = New RangeInMinExMax: RangeInMinExMax.New_ MinValue, MaxValue
End Function
Public Function RangeInMinInMax(ByVal MinValue As Double, ByVal MaxValue As Double) As RangeInMinInMax
    Set RangeInMinInMax = New RangeInMinInMax: RangeInMinInMax.New_ MinValue, MaxValue
End Function

Public Function IRange(ByVal MinValue As Double, ByVal MaxValue As Double, ByVal RangeType As ERangeType) As IRange
    Select Case RangeType
    Case ERangeType.ExMinExMax: Set IRange = MNew.RangeExMinExMax(MinValue, MaxValue)
    Case ERangeType.ExMinInMax: Set IRange = MNew.RangeExMinInMax(MinValue, MaxValue)
    Case ERangeType.InMinExMax: Set IRange = MNew.RangeInMinExMax(MinValue, MaxValue)
    Case ERangeType.InMinInMax: Set IRange = MNew.RangeInMinInMax(MinValue, MaxValue)
    End Select
End Function

Public Function FormatRange(ByVal vt As VbVarType, ByVal Min As Double, ByVal Max As Double) As String
    Dim s As String
    Select Case vt
    Case VbVarType.vbDate
        Dim datMin As Date: datMin = Min
        Dim datMax As Date: datMax = Max
        s = datMin & " | " & datMax
    Case Else
        s = Min & " | " & Max
    End Select
    FormatRange = s
End Function

