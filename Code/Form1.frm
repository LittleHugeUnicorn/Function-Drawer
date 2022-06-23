VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ãèµã·¨×÷º¯ÊýÍ¼Ïñ"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8265
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   6840
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   6480
      Width           =   3495
   End
   Begin VB.PictureBox Picoutput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   120
      ScaleHeight     =   6
      ScaleMode       =   0  'User
      ScaleWidth      =   8
      TabIndex        =   2
      Top             =   120
      Width           =   8000
      Begin VB.Line linemousex 
         BorderColor     =   &H00808080&
         X1              =   1.205
         X2              =   1.687
         Y1              =   0.724
         Y2              =   0.724
      End
      Begin VB.Line linemousey 
         BorderColor     =   &H00808080&
         X1              =   0.844
         X2              =   0.844
         Y1              =   0.482
         Y2              =   1.085
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ãèµã»æÍ¼"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   1
      Top             =   6120
      Width           =   1765
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "H(x)="
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "G(x)="
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F(x)="
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fx As String
Dim gx As String
Dim hx As String

Private Sub Command1_Click()
Picoutput.Cls
fx = Text1.Text
gx = Text2.Text
hx = Text3.Text
chushi
If Replace(fx, " ", "") <> "" Then Call Draw(fx, vbRed)
If Replace(gx, " ", "") <> "" Then Call Draw(gx, vbMagenta)
If Replace(hx, " ", "") <> "" Then Call Draw(hx, vbBlue)
End Sub

Private Sub Form_Load()
Form1.Show
chushi
End Sub

Function JS(ByVal Expressions As String) As String
   Dim Mssc As Object
   Set Mssc = CreateObject("MSScriptControl.ScriptControl")
   Mssc.Language = "vbscript"
   On Error GoTo EvalErr
   JS = Mssc.Eval(Expressions)
   Exit Function
EvalErr:
JS = "Error "
End Function

Sub chushi()
Dim i, j As Integer
Show
Picoutput.DrawWidth = 1
    For i = 1 To 7
    Picoutput.Line (i, 0)-(i, 6), RGB(200, 200, 200)
    Next i
    For j = 1 To 6
    Picoutput.Line (0, j)-(8, j), RGB(200, 200, 200)
    Next j
Picoutput.DrawWidth = 2
Picoutput.Line (4, 0)-(4, 6), RGB(0, 0, 0)
Picoutput.Line (0, 3)-(8, 3), RGB(0, 0, 0)
Picoutput.CurrentX = 0
Picoutput.CurrentY = 0
Picoutput.ForeColor = vbRed
Picoutput.Print "F(x)= " & fx
Picoutput.ForeColor = vbMagenta
Picoutput.Print "G(x)= " & gx
Picoutput.ForeColor = vbBlue
Picoutput.Print "H(x)= " & hx
Picoutput.ForeColor = vbBlack
Picoutput.CurrentX = 4.1
Picoutput.CurrentY = 0
Picoutput.Print "y"
Picoutput.CurrentX = 7.8
Picoutput.CurrentY = 3
Picoutput.Print "x"
End Sub

Private Sub Draw(ByRef hs As String, ByRef color As OLE_COLOR)
Dim dr As String
Dim i As Integer
chushi
    For i = -400 To 400
    On Error Resume Next
    dr = Replace(Trim(hs), "x", Str(i / 100))
    'Form1.Print JS(dr) 'µ÷ÊÔÓÃ
    'Picoutput.PSet ((i / 100 + 4), Val(-JS(dr)) + 3), vbRed
    Picoutput.Circle ((i / 100 + 4), Val(-JS(dr)) + 3), 0.015, color
    Next i
End Sub


Private Sub Picoutput_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
linemousex.X1 = 0
linemousex.X2 = Picoutput.Width
linemousex.Y1 = Y
linemousex.Y2 = Y
linemousey.Y1 = 0
linemousey.Y2 = Picoutput.Height
linemousey.X1 = X
linemousey.X2 = X
End Sub
