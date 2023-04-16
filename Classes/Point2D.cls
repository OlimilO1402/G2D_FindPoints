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
Private m_X   As Double
Private m_Y   As Double
Private m_Z   As Double
Private m_Tag As String

Friend Sub New_(ByVal aX As Double, ByVal aY As Double, ByVal aZ As Double, ByVal aTag As String)
    m_X = aX: m_Y = aY: m_Z = aZ: m_Tag = aTag
End Sub

Friend Sub NewC(other As Point3D)
    m_X = other.X: m_Y = other.Y: m_Z = other.Z: m_Tag = other.Tag
End Sub

Friend Function Clone() As Point3D
    Set Clone = New Point3D: Clone.NewC Me
End Function

Public Property Get Key() As String
    Key = Format(m_X, "0.###") & " " & Format(m_Y, "0.###") & " " & Format(m_Z, "0.###")
End Property

Public Function ToStr() As String
    ToStr = m_Tag & " {X=" & Format(m_X, "0.0##") & "; Y=" & Format(m_Y, "0.0##") & "; Z=" & Format(m_Z, "0.0##") & "}"
End Function

Public Function Compare(other As Point3D) As Long
Try: On Error GoTo Catch
    Dim i0 As Long: i0 = CLng(m_Tag)
    Dim i1 As Long: i1 = CLng(other.Tag)
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

Public Property Get X() As Double
    X = m_X
End Property
Public Property Let X(ByVal Value As Double)
    m_X = Value
End Property

Public Property Get Y() As Double
    Y = m_Y
End Property
Public Property Let Y(ByVal Value As Double)
    m_Y = Value
End Property

Public Property Get Z() As Double
    Z = m_Z
End Property
Public Property Let Z(ByVal Value As Double)
    m_Z = Value
End Property

Public Property Get Tag() As String
    Tag = m_Tag
End Property
Public Property Let Tag(ByVal Value As String)
    m_Tag = Value
End Property

Public Sub AddTag(ByVal aNewTag As String)
    m_Tag = m_Tag & " " & aNewTag
End Sub

Public Function DistanceTo(X As Double, Y As Double) As Double
    Dim dX As Double: dX = m_X - X
    Dim dY As Double: dY = m_Y - Y
    DistanceTo = VBA.Math.Sqr(dX * dX + dY * dY)
End Function

Public Function IsSame(other As Point3D) As Boolean
    IsSame = ObjPtr(Me) = ObjPtr(other)
End Function

Public Function Equals(other As Point3D) As Boolean
    Equals = m_X = other.X: If Not Equals Then Exit Function
    Equals = m_Y = other.Y: If Not Equals Then Exit Function
    Equals = m_Z = other.Z
End Function
