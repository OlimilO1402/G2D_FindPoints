VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Tag As String 'Tag, Bezeichnung, Name ...
Private m_X   As Double
Private m_Y   As Double
Private m_Z   As Double

Friend Sub New_(TagName As String, x As Double, y As Double, z As Double)
    m_Tag = TagName: m_X = x: m_Y = y: m_Z = z
End Sub

Public Sub NewC(other As Point3D)
    m_X = other.x: m_Y = other.y: m_Z = other.z: m_Tag = other.Tag
End Sub

Public Function Clone() As Point3D
    Set Clone = New Point3D: Clone.NewC Me
End Function

Public Function SetParams(Optional TagName, Optional CoordX, Optional CoordY, Optional CoordZ)
    If Not IsMissing(TagName) Then m_Tag = CStr(TagName)
    If Not IsMissing(CoordX) Then m_X = CDbl(CoordX)
    If Not IsMissing(CoordY) Then m_Y = CDbl(CoordY)
    If Not IsMissing(CoordZ) Then m_Z = CDbl(CoordZ)
End Function

Public Property Get Key() As String
    Key = Format(m_X, "0.###") & " " & Format(m_Y, "0.###") & " " & Format(m_Z, "0.###")
End Property

Public Function ToStr() As String
    ToStr = m_Tag & " {X=" & Format(m_X, "0.0##") & "; Y=" & Format(m_Y, "0.0##") & "; Z=" & Format(m_Z, "0.0##") & "}"
End Function

Public Function Compare(other As Point3D) As Long
Try: On Error GoTo Catch
    Dim i0 As Long: i0 = Me.TagToLong 'CLng(m_Tag)
    Dim i1 As Long: i1 = other.TagToLong 'CLng(other.Tag)
    If i0 > i1 Then
        Compare = 1
    ElseIf i0 < i1 Then
        Compare = -1
    End If
Catch:
End Function
Public Function ToTBCB() As String
    ToTBCB = m_Tag & vbTab & Format(m_X, "0.0##") & vbTab & Format(m_Y, "0.0##") & vbTab & Format(m_Z, "0.0##")
End Function

Public Property Get x() As Double
    x = m_X
End Property
'Public Property Let X(ByVal Value As Double)
'    m_X = Value
'End Property

Public Property Get y() As Double
    y = m_Y
End Property
'Public Property Let Y(ByVal Value As Double)
'    m_Y = Value
'End Property

Public Property Get z() As Double
    z = m_Z
End Property
'Public Property Let Z(ByVal Value As Double)
'    m_Z = Value
'End Property

Public Sub InvertX()
    m_X = -m_X
End Sub

Public Sub InvertY()
    m_Y = -m_Y
End Sub

Public Sub InvertZ()
    m_Z = -m_Z
End Sub

Public Property Get Tag() As String
    Tag = m_Tag
End Property
'Public Property Let Tag(ByVal Value As String)
'    m_Tag = Value
'End Property

Public Function TagToLong() As Long
    Dim s As String: s = m_Tag
    If Len(s) = 0 Then Exit Function
    Dim i As Long, c As String, si As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case AscW(c)
        Case 48 To 57: si = si & c
        End Select
    Next
    TagToLong = CLng(si)
End Function

Public Sub AddTag(ByVal aNewTag As String)
    m_Tag = m_Tag & " " & aNewTag
End Sub

Public Function DistanceTo(x As Double, y As Double) As Double
    Dim dX As Double: dX = m_X - x
    Dim dY As Double: dY = m_Y - y
    DistanceTo = VBA.Math.Sqr(dX * dX + dY * dY)
End Function

Public Function IsSame(other As Point3D) As Boolean
    IsSame = ObjPtr(Me) = ObjPtr(other)
End Function

Public Function Equals(other As Point3D) As Boolean
    Equals = m_X = other.x: If Not Equals Then Exit Function
    Equals = m_Y = other.y: If Not Equals Then Exit Function
    Equals = m_Z = other.z
End Function
