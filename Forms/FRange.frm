VERSION 5.00
Begin VB.Form FRange 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TBRangeMin 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TBRangeMax 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox ChkMinIncluded 
      Caption         =   "Minimum is included"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   2295
   End
   Begin VB.CheckBox ChkMaxIncluded 
      Caption         =   "Maximum is included"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   2295
   End
   Begin VB.TextBox TBCheckVal 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton BtnCheckInside 
      Caption         =   "Is Value inside Range?"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton BtnCheckOutside 
      Caption         =   "Is Value outside Range?"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label LblRangeMin 
      AutoSize        =   -1  'True
      Caption         =   "Minimum:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LblRangeMax 
      AutoSize        =   -1  'True
      Caption         =   "Maximum:"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Value to check:"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   1305
   End
End
Attribute VB_Name = "FRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'set default-values
    TBRangeMin.Text = Format(-2, "0.000")
    TBRangeMax.Text = Format(14, "0.000")
    TBCheckVal.Text = Format(5, "0.000")
End Sub

Private Sub TBRangeMin_LostFocus()
    TBRangeMin.Text = Format(GetRangeMin, "0.000")
End Sub

Private Sub TBRangeMax_LostFocus()
    TBRangeMax.Text = Format(GetRangeMax, "0.000")
End Sub

Private Sub TBCheckVal_LostFocus()
    TBCheckVal.Text = Format(GetCheckVal, "0.000")
End Sub

Private Function GetRangeMin() As Double
    Dim s As String: s = TBRangeMin.Text: If Len(s) = 0 Then s = "0"
    GetRangeMin = CDbl(s)
End Function

Private Function GetRangeMax() As Double
    Dim s As String: s = TBRangeMax.Text: If Len(s) = 0 Then s = "0"
    GetRangeMax = CDbl(s)
End Function

Private Function GetCheckVal() As Double
    Dim s As String: s = TBCheckVal.Text: If Len(s) = 0 Then s = "0"
    GetCheckVal = CDbl(s)
End Function

Private Function GetRange() As IRange
    Dim rmin As Double: rmin = GetRangeMin
    Dim rmax As Double: rmax = GetRangeMax
    If ChkMinIncluded.Value Then
        If ChkMaxIncluded.Value Then
            Set GetRange = MNew.RangeInMinInMax(rmin, rmax)
        Else
            Set GetRange = MNew.RangeInMinExMax(rmin, rmax)
        End If
    Else
        If ChkMaxIncluded.Value Then
            Set GetRange = MNew.RangeExMinInMax(rmin, rmax)
        Else
            Set GetRange = MNew.RangeExMinExMax(rmin, rmax)
        End If
    End If
End Function

Private Sub BtnCheckInside_Click()
    Dim Value As Double: Value = GetCheckVal
    Dim rng As IRange: Set rng = GetRange
    MsgBox "Value " & Value & " is" & IIf(Not rng.IsIn(Value), " not", "") & " inside the range " & rng.ToStr
End Sub

Private Sub BtnCheckOutside_Click()
    Dim Value As Double:   Value = GetCheckVal
    Dim rng   As IRange: Set rng = GetRange
    Dim ct    As Long:        ct = rng.IsOut(Value)
    Dim b     As Boolean:      b = CBool(ct)
    MsgBox "Value " & Value & " is" & IIf(Not b, " not", "") & " outside the range " & rng.ToStr & vbCrLf & IIf(b, Value & IIf(ct < 0, " is lower than the minimum", " exceeds the maximum"), "")
End Sub

'Proof that Date and Double are the same
'Private Type TDate
'    Value As Date
'End Type
'
'Private Type TDouble
'    Value As Double
'End Type
'
'Private Type TByte8
'    Value1 As Byte
'    Value2 As Byte
'    Value3 As Byte
'    Value4 As Byte
'    Value5 As Byte
'    Value6 As Byte
'    Value7 As Byte
'    Value8 As Byte
'End Type
'
'Private Function TByte8_Equals(this As TByte8, other As TByte8) As Boolean
'    With this
'        If .Value1 <> other.Value1 Then Exit Function
'        If .Value2 <> other.Value2 Then Exit Function
'        If .Value3 <> other.Value3 Then Exit Function
'        If .Value4 <> other.Value4 Then Exit Function
'        If .Value5 <> other.Value5 Then Exit Function
'        If .Value6 <> other.Value6 Then Exit Function
'        If .Value7 <> other.Value7 Then Exit Function
'        If .Value8 <> other.Value8 Then Exit Function
'    End With
'    TByte8_Equals = True
'End Function
'
'Private Sub Form_Load()
'    Dim dat As TDate:   dat.Value = Now
'    Dim dbl As TDouble: dbl.Value = dat.Value
'    Dim byt1 As TByte8: LSet byt1 = dat
'    Dim byt2 As TByte8: LSet byt2 = dbl
'    If TByte8_Equals(byt1, byt2) Then MsgBox "OK identisch"
'End Sub

