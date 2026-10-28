Attribute VB_Name = "MNew"
Option Explicit

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


