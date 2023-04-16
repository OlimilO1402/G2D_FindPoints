VERSION 5.00
Begin VB.Form FPoint3D 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtTag 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox TxtZ 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox TxtY 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox TxtX 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tag:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Z:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   165
   End
End
Attribute VB_Name = "FPoint3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Point As Point3D
Private m_Result As VbMsgBoxResult

Public Function ShowDialog(aPoint3D As Point3D, FOwner As Form) As VbMsgBoxResult
    Set m_Point = aPoint3D.Clone
    UpdateView
    SelectTag
    Me.Show vbModal, FOwner
    aPoint3D.NewC m_Point
    ShowDialog = m_Result
End Function

Public Sub UpdateView()
    TxtX.Text = Format(m_Point.X, "0.000")
    TxtY.Text = Format(m_Point.Y, "0.000")
    TxtZ.Text = Format(m_Point.Z, "0.000")
    TxtTag.Text = m_Point.Tag
End Sub

Sub SelectTag()
    'TxtTag.SelText = TxtTag.Text
    'TxtTag.SelStart = 0
    'TxtTag.SelLength = Len(TxtTag.Text)
    'TxtTag.SetFocus
    TxtTag.SelStart = 0
    TxtTag.SelLength = Len(TxtTag.Text)
    'MsgBox "OK selected"
End Sub

Private Sub BtnOK_Click()
    m_Result = vbOK
    GetallValues
    Unload Me
End Sub

Sub GetallValues()
    GetX
    GetY
    GetZ
    TxtTag_LostFocus
End Sub

Private Sub BtnCancel_Click()
    m_Result = vbCancel
    Unload Me
End Sub

Private Sub TxtX_LostFocus()
    GetX
    UpdateView
End Sub

Private Sub TxtY_LostFocus()
    GetY
    UpdateView
End Sub

Private Sub TxtZ_LostFocus()
    GetZ
    UpdateView
End Sub

Sub GetX()
    Dim Value As Double, s As String: s = TxtX.Text
    If Not MString.Double_TryParse(s, Value) Then
        MsgBox "Could not convert to Double :" & s
        Exit Sub
    End If
    m_Point.X = Value
End Sub

Sub GetY()
    Dim Value As Double, s As String: s = TxtY.Text
    If Not MString.Double_TryParse(s, Value) Then
        MsgBox "Could not convert to Double :" & s
        Exit Sub
    End If
    m_Point.Y = Value
End Sub

Sub GetZ()
    Dim Value As Double, s As String: s = TxtZ.Text
    If Not MString.Double_TryParse(s, Value) Then
        MsgBox "Could not convert to Double :" & s
        Exit Sub
    End If
    m_Point.Z = Value
End Sub

Private Sub TxtTag_LostFocus()
    m_Point.Tag = TxtTag.Text
End Sub

