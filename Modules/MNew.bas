Attribute VB_Name = "MNew"
Option Explicit

Public Function Point3D(ByVal aX As Double, ByVal aY As Double, ByVal aZ As Double, ByVal aTag As String) As Point3D
    Set Point3D = New Point3D: Point3D.New_ aX, aY, aZ, aTag
End Function

Public Function GraphicView(aCanvas As PictureBox, aPoint3DList As List) As GraphicView
    Set GraphicView = New GraphicView: GraphicView.New_ aCanvas, aPoint3DList
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

