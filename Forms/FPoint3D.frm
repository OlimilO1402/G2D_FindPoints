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
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
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
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
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
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
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
Private m_Point  As Point3D
'Private m_Result As VbMsgBoxResult
'in the Form we need:
'Private m_Obj As <AnyClass e.g. City>
'
'Public Sub UpdateView(Obj)
'    Set m_Obj = Obj
'    With m_Obj
'        TxtName.Text = .Name
'        TxtPostalCode.Text = .PostalCode
'    End With
'End Sub
'
'Public Function UpdateData(Obj) As Boolean
'    m_City.NewC MNew.City(TxtName.Text, TxtPostalCode.Text)
'    UpdateData = True
'End Function

Public Sub UpdateView(Obj)
    Set m_Point = Obj
    'every coord shows ponly the first 3 digits
    'if the tb gest the focus, all digits are shown
    With m_Point
        TxtTag.text = .Tag
        TxtX.text = Format(.x, "0.000")
        TxtY.text = Format(.y, "0.000")
        TxtZ.text = Format(.z, "0.000")
    End With
End Sub

Public Function UpdateData(Obj) As Boolean
'???
    UpdateData = True
End Function

'Public Function ShowDialog(aPoint3D As Point3D, FOwner As Form) As VbMsgBoxResult
'    Set m_Point = aPoint3D.Clone
'    UpdateView
'    SelectTag
'    Me.Show vbModal, FOwner
'    ShowDialog = m_Result
'    If ShowDialog = vbCancel Then Exit Function
'    'aPoint3D.NewC m_Point
'End Function

Sub SelectTag()
    'TxtTag.SelText = TxtTag.Text
    'TxtTag.SelStart = 0
    'TxtTag.SelLength = Len(TxtTag.Text)
    'TxtTag.SetFocus
    TxtTag.SelStart = 0
    TxtTag.SelLength = Len(TxtTag.text)
    'MsgBox "OK selected"
End Sub


'Private Sub BtnOK_Click()
'    m_R 'esult = vbOK
'    'GetallValues
'    Unload Me
'End Sub

'Sub GetallValues()
'    'GetX
'    'GetY
'    'GetZ
'    TxtTag_LostFocus
'End Sub

'Private Sub BtnCancel_Click()
'    m_Result = vbCancel
'    Unload Me
'End Sub

Private Sub TxtX_GotFocus()
    Dim st As String: st = TxtX.text
    Dim sp As String: sp = m_Point.x
    If Len(st) < Len(sp) Then TxtX.text = m_Point.x
End Sub
Private Sub TxtX_LostFocus()
    GetX
    UpdateView m_Point
End Sub

Private Sub TxtY_GotFocus()
    Dim st As String: st = TxtY.text
    Dim sp As String: sp = m_Point.y
    If Len(st) < Len(sp) Then TxtY.text = m_Point.y
End Sub
Private Sub TxtY_LostFocus()
    GetY
    UpdateView m_Point
End Sub

Private Sub TxtZ_GotFocus()
    Dim st As String: st = TxtZ.text
    Dim sp As String: sp = m_Point.z
    If Len(st) < Len(sp) Then TxtZ.text = m_Point.z
End Sub
Private Sub TxtZ_LostFocus()
    GetZ
    UpdateView m_Point
End Sub

Sub GetX()
    Dim Value As Double, s As String: s = TxtX.text
    If Not MString.Double_TryParse(s, Value) Then
        MsgBox "Could not convert to Double :" & s
        Exit Sub
    End If
    m_Point.SetParams CoordX:=Value
End Sub

Sub GetY()
    Dim Value As Double, s As String: s = TxtY.text
    If Not MString.Double_TryParse(s, Value) Then
        MsgBox "Could not convert to Double :" & s
        Exit Sub
    End If
    m_Point.SetParams CoordY:=Value
End Sub

Sub GetZ()
    Dim Value As Double, s As String: s = TxtZ.text
    If Not MString.Double_TryParse(s, Value) Then
        MsgBox "Could not convert to Double :" & s
        Exit Sub
    End If
    m_Point.SetParams CoordZ:=Value
End Sub

Private Sub TxtTag_LostFocus()
    m_Point.SetParams TagName:=TxtTag.text
End Sub
