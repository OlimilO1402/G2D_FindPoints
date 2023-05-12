VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Points As List '(Of Point3D)

Private Sub Class_Initialize()
    Set m_Points = MNew.List(vbObject)
End Sub

Public Function AddPoint(NewPoint As Point3D) As Point3D
    Set AddPoint = NewPoint
    m_Points.Add NewPoint
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Public Property Get Point(ByVal Index As Long) As Point3D
    Set Point = m_Points.Item(Index)
End Property

Public Property Get Points() As List
    Set Points = m_Points
End Property
