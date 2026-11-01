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
   Begin VB.CommandButton BtnCopyToClipBoard 
      Caption         =   "Bsp->ClipBoard"
      Height          =   375
      Left            =   9720
      TabIndex        =   10
      Top             =   0
      Width           =   1695
   End
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

'Copy Paste from Excel like e.g.:
'1   1,234567    1,234567    0
'2   1,234567    10  0
'3   10  10  0
'4   10  5   0
'5   12  5   0
'6   12  1,234567    0


Private Sub Form_Load()
    mnuPopUpPointList.Visible = False
    Set m_Points = MNew.List(vbObject, , True)
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub BtnCopyToClipBoard_Click()
    Dim s As String
    s = s & "1" & vbTab & "1,1234567" & vbTab & "1,234567" & vbTab & 0 & vbCrLf
    s = s & "2" & vbTab & "1,1234567" & vbTab & "10" & vbTab & 0 & vbCrLf
    s = s & "3" & vbTab & "10" & vbTab & "10" & vbTab & 0 & vbCrLf
    s = s & "4" & vbTab & "10" & vbTab & "5" & vbTab & 0 & vbCrLf
    s = s & "5" & vbTab & "12" & vbTab & "5" & vbTab & 0 & vbCrLf
    s = s & "6" & vbTab & "12" & vbTab & "1,234567" & vbTab & 0 & vbCrLf
    Clipboard.SetText s
End Sub

Private Sub Form_Resize()
    Dim L As Single, t As Single, W As Single, H As Single
    L = TxtData.Left: t = TxtData.Top: W = TxtData.Width: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then TxtData.Move L, t, W, H
    L = List1.Left: t = List1.Top: W = List1.Width: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then List1.Move L, t, W, H
    L = PBcanvas.Left: t = PBcanvas.Top: W = Me.ScaleWidth - L: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then PBcanvas.Move L, t, W, H
    PBcanvas.Cls
    UpdateGraphicView
End Sub

Private Sub BtnReadPoints_Click()
    If TxtData.text = vbNullString Then
        TxtData.text = Clipboard.GetText
    End If
    Dim s As String: s = MString.GetTabbedText(TxtData.text)
    'Debug.Print s
    ReadPoints s
    Set mGraphicView = MNew.GraphicView(PBcanvas, m_Points)
    UpdateView
End Sub

Private Sub BtnClear_Click()
    Set m_Points = MNew.List(vbObject, , True)
    Set mGraphicView = MNew.GraphicView(PBcanvas, m_Points)
    'm_Points.Clear
    TxtData.text = vbNullString
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
    TxtData.text = s
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
    Dim x As Double, y As Double, z As Double, Tag As String
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
                j = j + 1: If j <= u Then Double_TryParse sa(j), x
                j = j + 1: If j <= u Then Double_TryParse sa(j), y
                j = j + 1: If j <= u Then Double_TryParse sa(j), z
                Set p = MNew.Point3D(Tag, x, y, z)
                If m_Points.ContainsKey(p.Key) Then
                    Set p0 = m_Points.ItemByKey(p.Key)
                    p0.AddTag p.Tag
                Else
                    m_Points.Add p
                End If
                x = 0: y = 0: z = 0
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
    'If FPoint3D.ShowDialog(p, Me) = vbCancel Then Exit Sub
    If MNew.ModalDialog(FPoint3D, FPoint3D.BtnOK, FPoint3D.BtnCancel).ShowDialog(p, Me) = vbCancel Then Exit Sub
    List1.List(i) = p.ToStr
    UpdateGraphicView
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        DeletePoint List1.ListIndex
    End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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


'Example for using this class in a ButtonClick on a Form:
'Private Sub BtnCityAdd_Click()
'    Dim City As City: Set City = MNew.City("Musterstadt", "00000")
'    If MNew.ModalDialog(FCity, FCity.BtnOK, FCity.BtnCancel).ShowDialog(City, Me) = vbCancel Then Exit Sub
'    MData.Cities_AddLB LstCities, City
'End Sub


Private Sub mGraphicView_PointSelected(p As Point3D, ByVal Index As Long)
    'MsgBox aPoint.X & " " & aPoint.Y
    List1.ListIndex = Index
    FPoint3D.Move Me.Left + (Me.Width - FPoint3D.Width) / 2, Me.Top + (Me.Height - FPoint3D.Height) / 2
    
    'If FPoint3D.ShowDialog(p, Me) = vbCancel Then Exit Sub
    If MNew.ModalDialog(FPoint3D, FPoint3D.BtnOK, FPoint3D.BtnCancel).ShowDialog(p, Me) = vbCancel Then Exit Sub
    UpdateView
End Sub

Private Sub PBcanvas_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        DeletePoint List1.ListIndex
    End If
End Sub

