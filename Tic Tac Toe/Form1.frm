VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe  v1.00  -  By George"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Reset Score"
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Again"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Line ver3 
      Visible         =   0   'False
      X1              =   1875
      X2              =   1875
      Y1              =   1560
      Y2              =   2880
   End
   Begin VB.Line ver1 
      Visible         =   0   'False
      X1              =   660
      X2              =   660
      Y1              =   1560
      Y2              =   2880
   End
   Begin VB.Line hor3 
      Visible         =   0   'False
      X1              =   600
      X2              =   1920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line hor1 
      Visible         =   0   'False
      X1              =   600
      X2              =   1920
      Y1              =   1630
      Y2              =   1630
   End
   Begin VB.Line ver2 
      Visible         =   0   'False
      X1              =   1270
      X2              =   1270
      Y1              =   1560
      Y2              =   2880
   End
   Begin VB.Line hor2 
      Visible         =   0   'False
      X1              =   480
      X2              =   2040
      Y1              =   2235
      Y2              =   2235
   End
   Begin VB.Line dear 
      Visible         =   0   'False
      X1              =   600
      X2              =   1920
      Y1              =   2880
      Y2              =   1560
   End
   Begin VB.Line arde 
      Visible         =   0   'False
      X1              =   600
      X2              =   1920
      Y1              =   1560
      Y2              =   2880
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "George Papadopoulos"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00808080&
      BorderWidth     =   6
      X1              =   600
      X2              =   4320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0080FF80&
      X1              =   4800
      X2              =   4800
      Y1              =   1080
      Y2              =   3360
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      X1              =   2760
      X2              =   2760
      Y1              =   1080
      Y2              =   3360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FF80&
      X1              =   2640
      X2              =   4920
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0080FF80&
      X1              =   2640
      X2              =   4920
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0080FF80&
      X1              =   2640
      X2              =   4920
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FF80&
      X1              =   2640
      X2              =   4920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FF80&
      X1              =   2640
      X2              =   4920
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF80&
      Height          =   255
      Left            =   3960
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF80&
      Height          =   255
      Left            =   3960
      Top             =   1440
      Width           =   615
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080C0FF&
      X1              =   120
      X2              =   2400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080C0FF&
      X1              =   2280
      X2              =   2280
      Y1              =   1080
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080C0FF&
      X1              =   240
      X2              =   240
      Y1              =   1080
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080C0FF&
      X1              =   120
      X2              =   2400
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFF80&
      Height          =   255
      Left            =   3960
      Top             =   2400
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFF80&
      Height          =   255
      Left            =   3960
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Games :"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label games 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Tieds :"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label tied 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label pc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Computer :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label user 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " User :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Stats "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tic Tac Toe"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   720
      TabIndex        =   12
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label box3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label box3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label box2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label box2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label box3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label box1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label box2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label box1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label box1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, o, msg As String
Private Sub box1_Click(Index As Integer)
If box1(Index).Caption = x Or box1(Index).Caption = o Then Exit Sub
box1(Index).Caption = x
If checkend = True Then MsgBox msg: Exit Sub
Call play(1, Index + 1)
End Sub
Private Sub box2_Click(Index As Integer)
If box2(Index).Caption = x Or box2(Index).Caption = o Then Exit Sub
box2(Index).Caption = x
If checkend = True Then MsgBox msg: Exit Sub
Call play(2, Index + 1)
End Sub

Private Sub box3_Click(Index As Integer)
If box3(Index).Caption = x Or box3(Index).Caption = o Then Exit Sub
box3(Index).Caption = x
If checkend = True Then MsgBox msg: Exit Sub
Call play(3, Index + 1)
End Sub



Function isclear(ByVal box As Integer, ByVal row As Integer) As Boolean
'this function will check if there are any x or o in a label
Select Case box
  Case 1
    If box1(row - 1).Caption = x Or box1(row - 1).Caption = o Then isclear = False Else isclear = True
  Case 2
    If box2(row - 1).Caption = x Or box2(row - 1).Caption = o Then isclear = False Else isclear = True
  Case 3
    If box3(row - 1).Caption = x Or box3(row - 1).Caption = o Then isclear = False Else isclear = True
End Select

End Function

Function akries(ByVal a As String) As Boolean

'strategies to make our program a winner
If box2(1).Caption = "o" And box1(0).Caption = "" And box1(2).Caption = "" And isclear(1, 1) Then mark 1, 1: akries = True: Exit Function
If box2(1).Caption = "o" And box1(0).Caption = "" And box1(2).Caption = "o" And isclear(1, 1) Then mark 1, 1: akries = True: Exit Function
If box2(1).Caption = "o" And box1(0).Caption = "o" And box1(2).Caption = "" And isclear(1, 3) Then mark 1, 3: akries = True: Exit Function




If box3(1).Caption = x And box2(2).Caption = x Then
If isclear(3, 3) Then mark 3, 3: akries = True: Exit Function
End If

If box1(1).Caption = x And box2(0).Caption = x Then
If isclear(1, 1) Then mark 1, 1: akries = True: Exit Function
End If

If box3(1).Caption = x And box2(0).Caption = x Then
If isclear(3, 1) Then mark 3, 1: akries = True: Exit Function
End If

If box1(1).Caption = x And box2(2).Caption = x Then
If isclear(1, 3) Then mark 1, 3: akries = True: Exit Function
End If


If isx(3, 1) = True And isx(2, 3) = True And isclear(3, 3) Then mark 3, 3: akries = True: Exit Function
If isx(2, 1) = True And isx(3, 3) = True And isclear(3, 2) Then mark 3, 2: akries = True: Exit Function

If isclear(1, 1) Then mark 1, 1: akries = True: Exit Function
If isclear(1, 3) Then mark 1, 3: akries = True: Exit Function
If isclear(3, 1) Then mark 3, 1: akries = True: Exit Function
If isclear(3, 3) Then mark 3, 3: akries = True: Exit Function
End Function

Sub random()
If akries("hi") = True Then Exit Sub
Do
DoEvents
a = Int(Rnd * 3) + 1
a1 = Int(Rnd * 3) + 1
If isclear(a, a1) = False Then ok = 0 Else ok = 1
Loop Until ok = 1
mark a, a1
End Sub

Function userwon() As String


If box1(0).Caption = x And box1(1).Caption = x And box1(2).Caption = x Then userwon = "user": hor1.Visible = True: GoTo o
If box1(0).Caption = o And box1(1).Caption = o And box1(2).Caption = o Then userwon = "pc": hor1.Visible = True: GoTo o

If box2(0).Caption = x And box2(1).Caption = x And box2(2).Caption = x Then userwon = "user": hor2.Visible = True: GoTo o
If box2(0).Caption = o And box2(1).Caption = o And box2(2).Caption = o Then userwon = "pc": hor2.Visible = True: GoTo o

If box3(0).Caption = x And box3(1).Caption = x And box3(2).Caption = x Then userwon = "user": hor3.Visible = True: GoTo o
If box3(0).Caption = o And box3(1).Caption = o And box3(2).Caption = o Then userwon = "pc": hor3.Visible = True: GoTo o




'check vertical

If box1(0).Caption = x And box2(0).Caption = x And box3(0).Caption = x Then userwon = "user": ver1.Visible = True: GoTo o
If box1(0).Caption = o And box2(0).Caption = o And box3(0).Caption = o Then userwon = "pc": ver1.Visible = True: GoTo o


If box1(1).Caption = x And box2(1).Caption = x And box3(1).Caption = x Then userwon = "user": ver2.Visible = True: GoTo o
If box1(1).Caption = o And box2(1).Caption = o And box3(1).Caption = o Then userwon = "pc": ver2.Visible = True: GoTo o


If box1(2).Caption = x And box2(2).Caption = x And box3(2).Caption = x Then userwon = "user": ver3.Visible = True: GoTo o
If box1(2).Caption = o And box2(2).Caption = o And box3(2).Caption = o Then userwon = "pc": ver3.Visible = True: GoTo o



'elekse diagonia
If box1(0).Caption = x And box2(1).Caption = x And box3(2).Caption = x Then userwon = "user": arde.Visible = True: GoTo o
If box1(0).Caption = o And box2(1).Caption = o And box3(2).Caption = o Then userwon = "pc": arde.Visible = True: GoTo o


If box1(2).Caption = x And box2(1).Caption = x And box3(0).Caption = x Then userwon = "user": dear.Visible = True: GoTo o
If box1(2).Caption = o And box2(1).Caption = o And box3(0).Caption = o Then userwon = "pc": dear.Visible = True: GoTo o






For n = 1 To 3
For t = 1 To 3
If isclear(n, t) = True Then GoTo r
Next t
Next n
userwon = "tied"
GoTo o
r:
userwon = ""
o:
End Function


Function checkend() As Boolean
a = userwon

If a = "tied" Then
    Command1.Enabled = True
    Command1.Caption = "Try Again"
    msg = "Tied!": tied.Caption = Val(tied.Caption) + 1: games.Caption = Val(games.Caption) + 1
Else
    If a = "user" Then msg = "You Won!!!": Command1.Caption = "Play Again": Command1.Enabled = True: user.Caption = Val(user.Caption + 1): games.Caption = Val(games.Caption) + 1
    If a = "pc" Then msg = "Computer Won!!": Command1.Caption = "Try Again": Command1.Enabled = True: pc.Caption = Val(pc.Caption + 1): games.Caption = Val(games.Caption) + 1
End If

checkend = True
If a = "" Then
checkend = False
Else
box1(0).Enabled = False
box1(1).Enabled = False
box1(2).Enabled = False
box2(0).Enabled = False
box2(1).Enabled = False
box2(2).Enabled = False
box3(0).Enabled = False
box3(1).Enabled = False
box3(2).Enabled = False
End If
End Function


Sub mark(ByVal box As Integer, ByVal row As Integer)

'this sub will mark a label with a "o"
If isclear(box, row) = False Then Call random: Exit Sub

Select Case box
    Case 1
        box1(row - 1).Caption = "o"
    Case 2
        box2(row - 1).Caption = "o"
    Case 3
        box3(row - 1).Caption = "o"
End Select

If checkend = True Then MsgBox msg: Exit Sub
End Sub



Function isx(ByVal box As Integer, ByVal row As Integer) As Boolean
'this function will check if there are any x or o in a label
Select Case box
  Case 1
    If box1(row - 1).Caption = x Then isx = False Else isx = True
  Case 2
    If box2(row - 1).Caption = x Then isx = False Else isx = True
  Case 3
    If box3(row - 1).Caption = x Then isx = False Else isx = True
End Select

End Function



Function iso(ByVal box As Integer, ByVal row As Integer) As Boolean
'this function will check if there are any x or o in a label
Select Case box
  Case 1
    If box1(row - 1).Caption = o Then iso = False Else iso = True
  Case 2
    If box2(row - 1).Caption = o Then iso = False Else iso = True
  Case 3
    If box3(row - 1).Caption = o Then iso = False Else iso = True
End Select

End Function





Sub play(ByVal box As Integer, ByVal row As Integer)
'this is the whole programs brain..after user plays..this sub will deside what move our program should make


'special moves to prevent our program form losing!!
If box1(0).Caption = x And box1(1).Caption = x And isclear(1, 3) Then mark 1, 3: Exit Sub

If box2(2).Caption = x And box3(1).Caption = x And isclear(3, 2) Then mark 3, 2: Exit Sub

If box1(1).Caption = x And box2(1).Caption = x And isclear(3, 2) Then mark 3, 2: Exit Sub

If box2(1).Caption = x And box3(1).Caption = x And isclear(1, 2) Then mark 1, 2: Exit Sub


If box1(0).Caption = x And box1(2).Caption = x And isclear(1, 2) Then mark 1, 2: Exit Sub

If box1(0).Caption = x And box2(1).Caption = x And isclear(3, 3) Then mark 3, 3: Exit Sub


If box2(1).Caption = "x" And box2(2).Caption = "x" And isclear(2, 1) = True Then mark 2, 1: Exit Sub

If box3(0).Caption = "x" And box2(1).Caption = "x" And isclear(1, 3) = True Then mark 1, 3: Exit Sub

If box3(1).Caption = "o" And box3(2).Caption = "o" And isclear(3, 1) = True Then mark 3, 1: Exit Sub

If box1(0).Caption = "x" And box3(0).Caption = "" And isclear(3, 1) Then mark 3, 1: Exit Sub

If box1(0).Caption = "x" And box1(1).Caption = "x" And isclear(1, 3) Then mark 1, 3: Exit Sub


If box1(0).Caption = "x" And box3(0).Caption = "" And isclear(2, 2) Then mark 2, 2: Exit Sub

If box1(2).Caption = o And box3(2).Caption = o And isclear(2, 3) Then mark 2, 3: Exit Sub














'first it will check and click on the center because it is the best move to start
If isclear(2, 2) = True Then
    mark 2, 2
    Exit Sub
End If


If box2(1).Caption = x And box1(0).Caption = x And isclear(2, 1) Then
mark 2, 1
Exit Sub
End If


'check the 1st row
If box1(0).Caption = o And box1(1).Caption = o Then
If isclear(1, 3) = True Then mark 1, 3: Exit Sub
End If

If box1(0).Caption = o And box1(2).Caption = o Then
If isclear(1, 2) = True Then mark 1, 2: Exit Sub
End If

If box1(1).Caption = o And box1(2).Caption = o Then
If isclear(1, 1) = True Then mark 1, 1: Exit Sub
End If



'check the 2nd row
If box2(0).Caption = o And box2(1).Caption = o Then
If isclear(2, 3) = True Then mark 2, 3: Exit Sub
End If

If box2(0).Caption = o And box2(2).Caption = o Then
If isclear(2, 2) = True Then mark 2, 2: Exit Sub
End If

If box2(1).Caption = o And box2(2).Caption = o Then
If isclear(2, 1) = True Then mark 2, 1: Exit Sub
End If




'check the 3rd row
If box3(0).Caption = o And box3(1).Caption = o Then
If isclear(1, 3) = True Then mark 1, 3: Exit Sub
End If

If box3(0).Caption = o And box3(2).Caption = o Then
If isclear(1, 2) = True Then mark 1, 2: Exit Sub
End If

If box3(1).Caption = o And box3(2).Caption = o Then
If isclear(1, 1) = True Then mark 1, 1: Exit Sub
End If










'check vertical
'check 1st colum
If box1(0).Caption = o And box2(0).Caption = o Then
If isclear(3, 1) = True Then mark 3, 1: Exit Sub
End If

If box2(0).Caption = o And box3(0).Caption = o Then
If isclear(1, 1) = True Then mark 1, 1: Exit Sub
End If

If box1(0).Caption = o And box3(0).Caption = o Then
If isclear(2, 1) = True Then mark 2, 1: Exit Sub
End If


'check 2nd colume
If box1(1).Caption = o And box2(1).Caption = o Then
If isclear(3, 2) = True Then mark 3, 2: Exit Sub
End If

If box2(1).Caption = o And box3(1).Caption = o Then
If isclear(1, 2) = True Then mark 1, 2: Exit Sub
End If

If box1(1).Caption = o And box3(1).Caption = o Then
If isclear(2, 2) = True Then mark 2, 2: Exit Sub
End If




'check 3rd colume
If box1(2).Caption = o And box2(2).Caption = o Then
If isclear(3, 3) = True Then mark 3, 3: Exit Sub
End If

If box2(2).Caption = o And box3(2).Caption = o Then
If isclear(1, 3) = True Then mark 1, 3: Exit Sub
End If

If box1(2).Caption = o And box3(2).Caption = o Then
If isclear(2, 3) = True Then mark 2, 3: Exit Sub
End If




'elexei ta diagonia \
If box1(0).Caption = o And box3(2).Caption = o Then
If isclear(2, 2) Then mark 2, 2: Exit Sub
End If

If box1(0).Caption = o And box2(1).Caption = o Then
If isclear(3, 3) Then mark 3, 3: Exit Sub
End If

If box2(1).Caption = o And box3(2).Caption = o Then
If isclear(1, 1) Then mark 1, 1: Exit Sub
End If


'elexei diagonia /
If box1(2).Caption = o And box3(0).Caption = o Then
If isclear(2, 2) Then mark 2, 2: Exit Sub
End If

If box1(2).Caption = o And box2(1).Caption = o Then
If isclear(3, 3) Then mark 3, 3: Exit Sub
End If

If box2(1).Caption = o And box3(0).Caption = o Then
If isclear(1, 3) Then mark 1, 3: Exit Sub
End If















'defend diagonia \
If box1(0).Caption = x And box3(2).Caption = x Then
If isclear(2, 2) Then mark 2, 2: Exit Sub
End If

If box1(0).Caption = x And box2(1).Caption = x Then
If isclear(3, 3) Then mark 3, 3: Exit Sub
End If

If box2(1).Caption = x And box3(2).Caption = x Then
If isclear(1, 1) Then mark 1, 1: Exit Sub
End If





'defend diagonia /
If box1(2).Caption = x And box3(0).Caption = x Then
If isclear(2, 2) Then mark 2, 2: Exit Sub
End If

If box1(2).Caption = x And box2(1).Caption = x Then
If isclear(3, 1) Then mark 3, 1: Exit Sub
End If

If box2(1).Caption = x And box3(0).Caption = x Then
If isclear(1, 3) Then mark 1, 3: Exit Sub
End If








'check the row if has already 2 "x" in line
Select Case row
Case 1
    If isx(box, 2) = False Then mark box, 3: GoTo Last
    If isx(box, 3) = False Then mark box, 2: GoTo Last
Case 2
    If isx(box, 1) = False Then mark box, 3: GoTo Last
    If isx(box, 3) = False Then mark box, 1: GoTo Last
Case 3
    If isx(box, 1) = False Then mark box, 2: GoTo Last
    If isx(box, 2) = False Then mark box, 1: GoTo Last
End Select





Select Case box
Case 1
    If isx(3, row) = False Then mark 2, row: GoTo Last
    If isx(2, row) = False Then mark 3, row: GoTo Last
Case 2
    If isx(1, row) = False Then mark 3, row: GoTo Last
    If isx(3, row) = False Then mark 1, row: GoTo Last
Case 3
    If isx(2, row) = False Then mark 1, row: GoTo Last
    If isx(1, row) = False Then mark 2, row: GoTo Last
End Select




















Call random

'Last Command To Exesute before Letting user To Play Again
Last:

End Sub















Private Sub Command1_Click()
'clear all label so we can play again
Command1.Enabled = False
msg = ""
hor1.Visible = False
hor2.Visible = False
hor3.Visible = False
ver1.Visible = False
ver2.Visible = False
ver3.Visible = False
dear.Visible = False
arde.Visible = False



box1(0).Enabled = True
box1(1).Enabled = True
box1(2).Enabled = True
box2(0).Enabled = True
box2(1).Enabled = True
box2(2).Enabled = True
box3(0).Enabled = True
box3(1).Enabled = True
box3(2).Enabled = True
box1(0).Caption = ""
box1(1).Caption = ""
box1(2).Caption = ""
box2(0).Caption = ""
box2(1).Caption = ""
box2(2).Caption = ""
box3(0).Caption = ""
box3(1).Caption = ""
box3(2).Caption = ""

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
MsgBox "         Tic Tac Toe  v1.00" & vbCrLf & "                     By" & vbCrLf & "       George Papadopoulos"
End Sub

Private Sub Command4_Click()
a = MsgBox("Are You Sure You Want To Clear Scores?", vbYesNo, "Clear Scores?")
If a = vbYes Then
tied.Caption = "0"
games.Caption = "0"
user.Caption = "0"
pc.Caption = "0"
Exit Sub
End If
End Sub

Private Sub Form_Load()
x = "x"
o = "o"
End Sub

