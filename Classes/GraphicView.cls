VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mCanvas As PictureBox
Attribute mCanvas.VB_VarHelpID = -1
Private m_Points  As List
Private m_MarginL As Double 'Left
Private m_MarginT As Double 'Top
Private m_MarginR As Double 'Right
Private m_MarginB As Double 'Bottom
Private m_Zoom    As Double '
'Private m_OriginX As Double
'Private m_OriginY As Double
Private m_MinX    As Double
Private m_MinY    As Double
Private m_MinZ    As Double
Private m_MaxX    As Double
Private m_MaxY    As Double
Private m_MaxZ    As Double

Private m_PointHighlighted As Point3D
Private m_PointSelected    As Point3D

Public Event MousePointInWorldCoords(ByVal PX As Double, ByVal PY As Double)
Public Event PointSelected(p As Point3D, ByVal Index As Long) 'for editing!

Private Sub Class_Initialize()
    m_MarginL = 1: m_MarginT = 1: m_MarginR = 1: m_MarginB = 1
End Sub

Friend Sub New_(aCanvas As PictureBox, aPoint3DList As List)
    Set mCanvas = aCanvas: Set m_Points = aPoint3DList
    mCanvas.ScaleMode = vbPixels
    CalcMinMaxExtents
End Sub

Public Property Get Canvas() As PictureBox
    Set Canvas = mCanvas
End Property

Public Property Get Points() As List
    Set Points = m_Points
End Property

Public Sub CalcMinMaxExtents()
    Dim X As Double, Y As Double, Z As Double
    Dim i As Long, p As Point3D
    For i = 0 To m_Points.Count - 1
        Set p = m_Points.Item(i)
        X = p.X: Y = p.Y: Z = p.Z
        m_MinX = Min(m_MinX, X): m_MaxX = Max(m_MaxX, X)
        m_MinY = Min(m_MinY, Y): m_MaxY = Max(m_MaxY, Y)
        m_MinZ = Min(m_MinZ, Z): m_MaxZ = Max(m_MaxZ, Z)
    Next
End Sub

'properties of the model-points
Public Property Get CenterX() As Double
    CenterX = m_MinX + (m_MaxX - m_MinX) / 2
End Property

Public Property Get CenterY() As Double
    CenterY = m_MinY + (m_MaxY - m_MinY) / 2
End Property

Public Property Get CenterZ() As Double
    CenterZ = (m_MaxZ - m_MinZ) / 2
End Property

Public Property Get SizeMaxX() As Double
    SizeMaxX = Abs(m_MaxX - m_MinX)
End Property

Public Property Get SizeMaxY() As Double
    SizeMaxY = Abs(m_MaxY - m_MinY)
End Property

Public Property Get SizeMaxZ() As Double
    SizeMaxZ = Abs(m_MaxZ - m_MinZ)
End Property

Public Sub DrawPointsXY()
    'draws all points inside the canvas
    'World-coordinates
    Dim maxsizeX As Double: maxsizeX = m_MarginL + Me.SizeMaxX + m_MarginR
    Dim maxsizeY As Double: maxsizeY = m_MarginT + Me.SizeMaxY + m_MarginB
    Dim X0 As Double: X0 = Me.CenterX
    Dim Y0 As Double: Y0 = Me.CenterY
    Dim uniX As Double: uniX = mCanvas.ScaleWidth / maxsizeX
    Dim uniY As Double: uniY = mCanvas.ScaleHeight / maxsizeY
    m_Zoom = Min(uniX, uniY)
    'picture-coordinates
    Dim X As Single, XN As Single: XN = mCanvas.ScaleWidth / 2
    Dim Y As Single, YN As Single: YN = mCanvas.ScaleHeight / 2
    Dim i As Long, p As Point3D
    For i = 0 To m_Points.Count - 1
        Set p = m_Points.Item(i)
        X = XN + (p.X - X0) * m_Zoom
        Y = YN - (p.Y - Y0) * m_Zoom
        mCanvas.Circle (X, Y), 3, vbBlue
        mCanvas.Print p.Tag
    Next
End Sub

Public Sub HighlightPoint(p As Point3D)
    If Not m_PointHighlighted Is Nothing Then
        Point_DeleteHighlighting m_PointHighlighted
    End If
    Set m_PointHighlighted = p
    Point_DrawHighlighted p
End Sub

Private Sub Point_DrawHighlighted(p As Point3D)
    Dim X0 As Double: X0 = Me.CenterX
    Dim Y0 As Double: Y0 = Me.CenterY
    Dim XN As Single: XN = mCanvas.ScaleWidth / 2
    Dim YN As Single: YN = mCanvas.ScaleHeight / 2
    Dim X  As Single:  X = XN + (p.X - X0) * m_Zoom
    Dim Y  As Single:  Y = YN - (p.Y - Y0) * m_Zoom
    mCanvas.Circle (X, Y), 3, vbRed
End Sub
Private Sub Point_DeleteHighlighting(p As Point3D)
    Dim X0 As Double: X0 = Me.CenterX
    Dim Y0 As Double: Y0 = Me.CenterY
    Dim XN As Single: XN = mCanvas.ScaleWidth / 2
    Dim YN As Single: YN = mCanvas.ScaleHeight / 2
    Dim X  As Single:  X = XN + (p.X - X0) * m_Zoom
    Dim Y  As Single:  Y = YN - (p.Y - Y0) * m_Zoom
    mCanvas.Circle (X, Y), 3, vbBlue
End Sub

Private Sub mCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X0 As Double: X0 = Me.CenterX
    Dim Y0 As Double: Y0 = Me.CenterY
    Dim XN As Single: XN = mCanvas.ScaleWidth / 2
    Dim YN As Single: YN = mCanvas.ScaleHeight / 2
    Dim PX As Double: PX = X0 + (X - XN) / m_Zoom
    Dim PY As Double: PY = Y0 - (Y - YN) / m_Zoom
    Dim p As Point3D, i As Long
    If GetPointByCoords(PX, PY, p, i) Then
        If Not m_PointHighlighted Is Nothing Then
            If Not m_PointHighlighted.IsSame(p) Then
                Point_DeleteHighlighting m_PointHighlighted
            End If
        End If
        Set m_PointHighlighted = p
        Point_DrawHighlighted p
    Else
        If Not m_PointHighlighted Is Nothing Then
            Point_DeleteHighlighting m_PointHighlighted
        End If
    End If
    RaiseEvent MousePointInWorldCoords(PX, PY)
    
End Sub

Private Sub mCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> MouseButtonConstants.vbLeftButton Then Exit Sub
    'convert the picture-coorinates to world-coordinates
    Dim X0 As Double: X0 = Me.CenterX
    Dim Y0 As Double: Y0 = Me.CenterY
    Dim XN As Single: XN = mCanvas.ScaleWidth / 2
    Dim YN As Single: YN = mCanvas.ScaleHeight / 2
    Dim PX As Double: PX = X0 + (X - XN) / m_Zoom
    Dim PY As Double: PY = Y0 - (Y - YN) / m_Zoom
    Dim p As Point3D, i As Long
    If GetPointByCoords(PX, PY, p, i) Then
        RaiseEvent PointSelected(p, i)
    End If
End Sub

Private Sub mCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Function GetPointByCoords(X As Double, Y As Double, p_out As Point3D, i_out As Long) As Boolean
    Dim i As Long, p As Point3D
    For i = 0 To m_Points.Count - 1
        Set p = m_Points.Item(i)
        If p.DistanceTo(X, Y) < 0.2 Then
            Set p_out = p
            i_out = i
            GetPointByCoords = True
            Exit Function
        End If
    Next
End Function
