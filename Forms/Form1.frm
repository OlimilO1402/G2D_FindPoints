VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FindPoints"
   ClientHeight    =   10455
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   19815
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   19815
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnMirrorXatYaxis 
      Caption         =   "Mirror X@Y-axis <>"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton BtnMirrorYatXaxis 
      Caption         =   "Mirror Y@X-axis^v"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton BtnWriteToTBCB 
      Caption         =   "Write to TB&&CB"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton BtnSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton BtnReadPoints 
      Caption         =   "Read Points"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox PBcanvas 
      BackColor       =   &H80000005&
      DrawStyle       =   6  'Innen ausgefüllt
      Height          =   9975
      Left            =   6240
      ScaleHeight     =   9915
      ScaleWidth      =   13515
      TabIndex        =   3
      Top             =   360
      Width           =   13575
   End
   Begin VB.TextBox TxtData 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9960
      ItemData        =   "Form1.frx":1782
      Left            =   3120
      List            =   "Form1.frx":1784
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label LblMouseInWorldCoords 
      AutoSize        =   -1  'True
      Caption         =   "X: ; Y: ;"
      Height          =   195
      Left            =   11880
      TabIndex        =   5
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuPopUpPointList 
      Caption         =   "mnuPopUpPointList"
      Begin VB.Menu mnuPointsDeletePoint 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'copy-paste some point data into the textbox
'parse it, create all points objects and show the points in the view
'now select the points with your mouse and give it the right numbers
Private m_Points As List
Private WithEvents mGraphicView As GraphicView
Attribute mGraphicView.VB_VarHelpID = -1

Private Sub Form_Load()
    mnuPopUpPointList.Visible = False
    Set m_Points = MNew.List(vbObject, , True)
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    L = TxtData.Left: T = TxtData.Top: W = TxtData.Width: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then TxtData.Move L, T, W, H
    L = List1.Left: T = List1.Top: W = List1.Width: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
    L = PBcanvas.Left: T = PBcanvas.Top: W = Me.ScaleWidth - L: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then PBcanvas.Move L, T, W, H
    PBcanvas.Cls
    UpdateGraphicView
End Sub

Private Sub BtnReadPoints_Click()
    If TxtData.Text = vbNullString Then
        TxtData.Text = Clipboard.GetText
    End If
    Dim s As String: s = MString.GetTabbedText(TxtData.Text)
    'Debug.Print s
    ReadPoints s
    Set mGraphicView = MNew.GraphicView(PBcanvas, m_Points)
    UpdateView
End Sub

Private Sub BtnClear_Click()
    Set m_Points = MNew.List(vbObject, , True)
    Set mGraphicView = MNew.GraphicView(PBcanvas, m_Points)
    'm_Points.Clear
    TxtData.Text = vbNullString
    UpdateView
End Sub

Private Sub BtnSort_Click()
    m_Points.Sort
    UpdateView
End Sub

Private Sub BtnWriteToTBCB_Click()
    Dim s As String
    Dim i As Long, p As Point3D
    For i = 0 To m_Points.Count - 1
        Set p = m_Points.Item(i)
        s = s & p.ToTBCB & vbCrLf
    Next
    TxtData.Text = s
    Clipboard.Clear
    Clipboard.SetText s
End Sub

Private Sub BtnMirrorYatXaxis_Click()
    'Flip all points vertically, the points above will become the points below and vice versa
    'Mirror ^v
    Dim i As Long, p As Point3D
    For i = 0 To m_Points.Count - 1
        Set p = m_Points.Item(i)
        p.InvertY
    Next
    'mGraphicView.CalcMinMaxExtents
    Set mGraphicView = MNew.GraphicView(PBcanvas, m_Points)
    UpdateView
End Sub

Private Sub BtnMirrorXatYaxis_Click()
    'Flip all points horizontally, the points left will become the points right and vice versa
    'Mirror <>
    Dim i As Long, p As Point3D
    For i = 0 To m_Points.Count - 1
        Set p = m_Points.Item(i)
        p.InvertX
    Next
    'mGraphicView.CalcMinMaxExtents
    Set mGraphicView = MNew.GraphicView(PBcanvas, m_Points)
    UpdateView
End Sub

Private Sub ReadPoints(s As String)
    Dim lines() As String: lines = Split(s, vbCrLf)
    Dim line As String, sa() As String
    Dim i As Long, j As Long, u As Long
    Dim X As Double, Y As Double, Z As Double, Tag As String
    Dim p As Point3D, p0 As Point3D
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If Len(line) Then
            sa = Split(line, vbTab)
'            If sa(0) = "K45" Then
'                Debug.Assert True
'            End If
            u = UBound(sa)
            'if only tag but no point -> p(0,0)+addedtags -> edit tag of p(0,0) afterwards
            If j <= u Then Tag = sa(j)
'                If Left(sa(j), 1) = "-" Then
'                    Debug.Assert True
'                ElseIf Left(sa(j + 1), 1) = "-" Then
'                    Debug.Assert True
'                ElseIf Left(sa(j + 2), 1) = "-" Then
'                    Debug.Assert True
'                ElseIf Left(sa(j + 2), 1) = "-" Then
'                    Debug.Assert True
'                End If
            'If j <> u Then
                j = j + 1: If j <= u Then Double_TryParse sa(j), X
                j = j + 1: If j <= u Then Double_TryParse sa(j), Y
                j = j + 1: If j <= u Then Double_TryParse sa(j), Z
                Set p = MNew.Point3D(X, Y, Z, Tag)
                If m_Points.ContainsKey(p.Key) Then
                    Set p0 = m_Points.ItemByKey(p.Key)
                    p0.AddTag p.Tag
                Else
                    m_Points.Add p
                End If
                X = 0: Y = 0: Z = 0
            'End If
            j = 0
        End If
    Next
End Sub

Sub UpdateView()
    'list all Points
    m_Points.ToListbox List1
    'now we got to draw all points
    PBcanvas.Cls
    mGraphicView.DrawPointsXY

End Sub

Sub UpdateGraphicView()
    PBcanvas.Cls
    If Not mGraphicView Is Nothing Then mGraphicView.DrawPointsXY
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Dim p As Point3D: Set p = m_Points.Item(i)
    If p Is Nothing Then Exit Sub
    mGraphicView.HighlightPoint p
End Sub

Private Sub List1_DblClick()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Dim p As Point3D: Set p = m_Points.Item(i)
    If p Is Nothing Then Exit Sub
    FPoint3D.Move Me.Left + (Me.Width - FPoint3D.Width) / 2, Me.Top + (Me.Height - FPoint3D.Height) / 2
    If FPoint3D.ShowDialog(p, Me) = vbCancel Then Exit Sub
    List1.List(i) = p.ToStr
    UpdateGraphicView
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        DeletePoint List1.ListIndex
    End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuPopUpPointList
    'ElseIf Button = MouseButtonConstants.vbLeftButton Then
    End If
End Sub

Private Sub mnuPointsDeletePoint_Click()
    DeletePoint List1.ListIndex
End Sub

Private Function DeletePoint(ByVal Index As Long)
    Dim p As Point3D: Set p = m_Points.Item(Index)
    Dim mr As VbMsgBoxResult: mr = MsgBox("Are you sure you want to delete this point? " & vbCrLf & p.ToStr, vbOKCancel)
    If mr = vbCancel Then
        Exit Function
    End If
    DeletePoint = True
    m_Points.Remove Index
    UpdateView
End Function

Private Sub mGraphicView_MousePointInWorldCoords(ByVal PX As Double, ByVal PY As Double)
    LblMouseInWorldCoords.Caption = "X: " & Format(PX, "0.00") & "; Y: " & Format(PY, "0.00")
End Sub

Private Sub mGraphicView_PointSelected(p As Point3D, ByVal Index As Long)
    'MsgBox aPoint.X & " " & aPoint.Y
    List1.ListIndex = Index
    FPoint3D.Move Me.Left + (Me.Width - FPoint3D.Width) / 2, Me.Top + (Me.Height - FPoint3D.Height) / 2
    If FPoint3D.ShowDialog(p, Me) = vbCancel Then Exit Sub
    UpdateView
End Sub

Private Sub PBcanvas_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        DeletePoint List1.ListIndex
    End If
End Sub

