Attribute VB_Name = "MNew"
Option Explicit

Public Function Point3D(ByVal TagName As String, ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Point3D
    Set Point3D = New Point3D: Point3D.New_ TagName, X, Y, Z
End Function

Public Function GraphicView(aCanvas As PictureBox, aPoint3DList As List) As GraphicView
    Set GraphicView = New GraphicView: GraphicView.New_ aCanvas, aPoint3DList
End Function

Public Function ModalDialog(Dialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
    Set ModalDialog = New ModalDialog: ModalDialog.New_ Dialog, BtnOK, BtnCancel
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

