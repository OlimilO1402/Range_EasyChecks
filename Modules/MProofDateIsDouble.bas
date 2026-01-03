Attribute VB_Name = "MProof"
Option Explicit

Private Type TDate
    Value As Date
End Type

Private Type TDouble
    Value As Double
End Type

Private Type TByte8
    Value1 As Byte
    Value2 As Byte
    Value3 As Byte
    Value4 As Byte
    Value5 As Byte
    Value6 As Byte
    Value7 As Byte
    Value8 As Byte
End Type

Private Function TByte8_Equals(this As TByte8, other As TByte8) As Boolean
    With this
        If .Value1 <> other.Value1 Then Exit Function
        If .Value2 <> other.Value2 Then Exit Function
        If .Value3 <> other.Value3 Then Exit Function
        If .Value4 <> other.Value4 Then Exit Function
        If .Value5 <> other.Value5 Then Exit Function
        If .Value6 <> other.Value6 Then Exit Function
        If .Value7 <> other.Value7 Then Exit Function
        If .Value8 <> other.Value8 Then Exit Function
    End With
    TByte8_Equals = True
End Function

Public Sub DateIsDouble()
    Dim dat As TDate:   dat.Value = Now
    Dim dbl As TDouble: dbl.Value = dat.Value
    Dim byt1 As TByte8: LSet byt1 = dat
    Dim byt2 As TByte8: LSet byt2 = dbl
    If TByte8_Equals(byt1, byt2) Then MsgBox "Date and Double are identical"
End Sub

