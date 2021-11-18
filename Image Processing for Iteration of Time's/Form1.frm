VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12480
   BeginProperty Font 
      Name            =   "Tw Cen MT Condensed Extra Bold"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   4
      Top             =   1320
      Width           =   3045
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   8400
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   3
      Top             =   5160
      Width           =   3060
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   4320
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   2
      Top             =   5160
      Width           =   3060
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "&Processing"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   480
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   5160
      Width           =   3060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Image Processing for Iteration of Time's"
      BeginProperty Font 
         Name            =   "unbutton"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   660
      Left            =   1320
      TabIndex        =   9
      Top             =   0
      Width           =   10140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Final Output Layer"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8400
      TabIndex        =   8
      Top             =   4680
      Width           =   2730
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Processing Layer"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4320
      TabIndex        =   7
      Top             =   4680
      Width           =   2505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Input Layer"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   6
      Top             =   4680
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Original Picture"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, col, col1
Dim n1, n2
Dim start As Boolean
Dim maxerrone As Double
Dim maxerrtwo As Double

Private Sub inplayer()
X = 0: Y = 0
For i = 0 To Picture4.ScaleHeight - 1
For j = 0 To Picture4.ScaleWidth - 1
col = Picture4.Point(X, Y) Mod 256
col1 = col / 255
ilayer.val(i, j) = col1
Picture1.PSet (X, Y), RGB(ilayer.val(i, j) * 255, ilayer.val(i, j) * 255, ilayer.val(i, j) * 255)
X = X + 1
Next
X = 0: Y = Y + 1
Next
End Sub

Private Sub hidlayer()
X = 0: Y = 0
For i = 0 To Picture1.ScaleHeight - 1
For j = 0 To Picture1.ScaleWidth - 1

If i = 0 And j = 0 Then
hlayer.val(i, j) = ilayer.val(i, j + 1) * hlayer.wt(1, 2) + ilayer.val(i + 1, j) * hlayer.wt(2, 1) + ilayer.val(i + 1, j + 1) * hlayer.wt(2, 2)
ElseIf i = 0 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
hlayer.val(i, j) = ilayer.val(i, j + 1) * hlayer.wt(1, 2) + ilayer.val(i, j - 1) * hlayer.wt(1, 0) + ilayer.val(i + 1, j) * hlayer.wt(2, 1) + ilayer.val(i + 1, j + 1) * hlayer.wt(2, 2) + ilayer.val(i + 1, j - 1) * hlayer.wt(2, 0)
ElseIf i = 0 And j = Picture1.ScaleWidth - 1 Then
hlayer.val(i, j) = ilayer.val(i, j - 1) * hlayer.wt(1, 0) + ilayer.val(i + 1, j - 1) * hlayer.wt(2, 0) + ilayer.val(i + 1, j) * hlayer.wt(2, 1)
ElseIf i = Picture1.ScaleHeight - 1 And j = 0 Then
hlayer.val(i, j) = ilayer.val(i, j + 1) * hlayer.wt(1, 2) + ilayer.val(i - 1, j) * hlayer.wt(0, 1) + ilayer.val(i - 1, j + 1) * hlayer.wt(0, 2)
ElseIf i = Picture1.ScaleHeight - 1 And j = Picture1.ScaleWidth - 1 Then
hlayer.val(i, j) = ilayer.val(i, j - 1) * hlayer.wt(1, 0) + ilayer.val(i - 1, j) * hlayer.wt(0, 1) + ilayer.val(i - 1, j - 1) * hlayer.wt(0, 0)
ElseIf i = Picture1.ScaleHeight - 1 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
hlayer.val(i, j) = ilayer.val(i, j + 1) * hlayer.wt(1, 2) + ilayer.val(i, j - 1) * hlayer.wt(1, 0) + ilayer.val(i - 1, j) * hlayer.wt(0, 1) + ilayer.val(i - 1, j + 1) * hlayer.wt(0, 2) + ilayer.val(i - 1, j - 1) * hlayer.wt(0, 0)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = 0 Then
hlayer.val(i, j) = ilayer.val(i, j + 1) * hlayer.wt(1, 2) + ilayer.val(i - 1, j) * hlayer.wt(0, 1) + ilayer.val(i - 1, j + 1) * hlayer.wt(0, 2) + ilayer.val(i + 1, j) * hlayer.wt(2, 1) + ilayer.val(i + 1, j + 1) * hlayer.wt(2, 2)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = Picture1.ScaleWidth - 1 Then
hlayer.val(i, j) = ilayer.val(i, j - 1) * hlayer.wt(1, 0) + ilayer.val(i - 1, j) * hlayer.wt(0, 1) + ilayer.val(i - 1, j - 1) * hlayer.wt(0, 0) + ilayer.val(i + 1, j) * hlayer.wt(2, 1) + ilayer.val(i + 1, j - 1) * hlayer.wt(2, 0)
Else
hlayer.val(i, j) = ilayer.val(i, j - 1) * hlayer.wt(1, 0) + ilayer.val(i, j + 1) * hlayer.wt(1, 2) + ilayer.val(i + 1, j) * hlayer.wt(2, 1) + ilayer.val(i + 1, j + 1) * hlayer.wt(2, 2) + ilayer.val(i + 1, j - 1) * hlayer.wt(2, 0) + ilayer.val(i - 1, j) * hlayer.wt(0, 1) + ilayer.val(i - 1, j + 1) * hlayer.wt(0, 2) + ilayer.val(i - 1, j - 1) * hlayer.wt(0, 0)
End If
hlayer.val(i, j) = 1 / (1 + Exp(-(hlayer.val(i, j) - 4)))
Picture2.PSet (X, Y), RGB(hlayer.val(i, j) * 255, hlayer.val(i, j) * 255, hlayer.val(i, j) * 255)
X = X + 1
Next
X = 0: Y = Y + 1
Next
End Sub
Private Sub outlayer()
X = 0: Y = 0
For i = 0 To Picture1.ScaleHeight - 1
For j = 0 To Picture1.ScaleWidth - 1

If i = 0 And j = 0 Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2)
ElseIf i = 0 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0)
ElseIf i = 0 And j = Picture1.ScaleWidth - 1 Then
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0) + hlayer.val(i + 1, j) * olayer.wt(2, 1)
ElseIf i = Picture1.ScaleHeight - 1 And j = 0 Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2)
ElseIf i = Picture1.ScaleHeight - 1 And j = Picture1.ScaleWidth - 1 Then
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0)
ElseIf i = Picture1.ScaleHeight - 1 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = 0 Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = Picture1.ScaleWidth - 1 Then
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0)
Else
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0)
End If
olayer.val(i, j) = 1 / (1 + Exp(-(olayer.val(i, j) - 4)))
Picture3.PSet (X, Y), RGB(olayer.val(i, j) * 255, olayer.val(i, j) * 255, olayer.val(i, j) * 255)
If (maxerrone = 0) Then
maxerrone = (1 + Exp(-(olayer.val(i, j) - 4)))
Else
If (maxerrone < (1 + Exp(-(olayer.val(i, j) - 4)))) Then
maxerrone = (1 + Exp(-(olayer.val(i, j) - 4)))
End If
End If
X = X + 1
Next
X = 0: Y = Y + 1
Next
End Sub

Private Sub Command1_Click()
start = True
If start Then
Dim e As Double
Dim count As Double
Dim diff As Double
Dim diffg As Double
maxerrone = 0
maxerrtwo = 0
diffg = 0.001
count = 0
e = 0
inplayer
Command1.Caption = "Step Number:" + CStr(count)
count = count + 1

Do While (1)
If (e = 0) Then
hidlayer
Command1.Caption = "Step Number:" + CStr(count)
count = count + 1
outlayer
Command1.Caption = "Step Number:" + CStr(count)
     count = count + 1

olayerTOhlayer_up_wt


hlayerTOilayer_up_wt

e = 1

    Else
    'else part
    olayerTOhidlayer
    Command1.Caption = "Step Number:" + CStr(count)
    count = count + 1
    hlayerTOolayer
    Command1.Caption = "Step Number:" + CStr(count)
    count = count + 1
    diff = 0
    diff = maxerrtwo - maxerrone
    maxerrone = maxerrtwo
    maxerrtwo = 0
    If diff <= diffg Then
      Command1.FontBold = True
      Command1.FontSize = 10
      Command1.Enabled = False
      Command1.Caption = " Stop "
      Label5.ForeColor = vbRed
      Label5.FontBold = True
      Label5.FontSize = 10
      Label5.Caption = " Result after " + CStr(count) + "th iteration "
      Exit Do
    End If
      'olayerTOhidlayer
    olayerTOhlayer_up_wt
    'hlayerTOilayer
    hlayerTOilayer_up_wt
    
   
End If
    Loop
Else
X = 0: Y = 0
For i = 0 To Picture1.ScaleHeight - 1
For j = 0 To Picture1.ScaleWidth - 1
ilayer.val(i, j) = olayer.val(i, j)
Next
X = 0: Y = Y + 1
Next
End If
End Sub


Private Sub hlayerTOilayer_up_wt()
Dim n As Double
n = (Picture1.ScaleHeight) * (Picture1.ScaleWidth)
Dim oj As Double
For k = 0 To Picture1.ScaleHeight - 1
For m = 0 To Picture1.ScaleWidth - 1
For i = 0 To 2
For j = 0 To 2
oj = hlayer.val(k, m)
If (oj >= 0 And oj <= 0.5) Then
hlayer.wt(i, j) = (ilayer.wt(i, j) + (0.001 * (-2 / n) * oj * (1 - oj) * olayer.val(k, m)))
ElseIf (oj >= 0.5 And oj <= 1) Then
hlayer.wt(i, j) = (ilayer.wt(i, j) + (0.001 * (2 / n) * oj * (1 - oj) * olayer.val(k, m)))
End If
Next
Next
Next
Next
End Sub
Private Sub olayerTOhlayer_up_wt()
Dim n As Double
n = (Picture1.ScaleHeight) * (Picture1.ScaleWidth)
Dim oj As Double
For k = 0 To Picture1.ScaleHeight - 1
For m = 0 To Picture1.ScaleWidth - 1
For i = 0 To 2
For j = 0 To 2
oj = olayer.val(k, m)
If (oj >= 0 And oj <= 0.5) Then
olayer.wt(i, j) = (hlayer.wt(i, j) + (0.001 * (-2 / n) * oj * (1 - oj) * hlayer.val(k, m)))
ElseIf (oj >= 0.5 And oj <= 1) Then
olayer.wt(i, j) = (hlayer.wt(i, j) + (0.001 * (2 / n) * oj * (1 - oj) * hlayer.val(k, m)))
End If
Next
Next
Next
Next
End Sub


Private Sub Form_Load()
start = False
n1 = Picture1.ScaleHeight: n2 = Picture1.ScaleWidth
ReDim Preserve ilayer.val(n1, n2)
ReDim Preserve hlayer.val(n1, n2)
ReDim Preserve olayer.val(n1, n2)
For i = 0 To 2
For j = 0 To 2
ilayer.wt(i, j) = 1
hlayer.wt(i, j) = 1
olayer.wt(i, j) = 1
Next
Next
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.ToolTipText = Picture1.Point(X, Y)
End Sub
Private Sub olayerTOhidlayer()
X = 0: Y = 0
For i = 0 To Picture1.ScaleHeight - 1
For j = 0 To Picture1.ScaleWidth - 1
If i = 0 And j = 0 Then
hlayer.val(i, j) = olayer.val(i, j + 1) * hlayer.wt(1, 2) + olayer.val(i + 1, j) * hlayer.wt(2, 1) + olayer.val(i + 1, j + 1) * hlayer.wt(2, 2)
ElseIf i = 0 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
hlayer.val(i, j) = olayer.val(i, j + 1) * hlayer.wt(1, 2) + olayer.val(i, j - 1) * hlayer.wt(1, 0) + olayer.val(i + 1, j) * hlayer.wt(2, 1) + olayer.val(i + 1, j + 1) * hlayer.wt(2, 2) + olayer.val(i + 1, j - 1) * hlayer.wt(2, 0)
ElseIf i = 0 And j = Picture1.ScaleWidth - 1 Then
hlayer.val(i, j) = olayer.val(i, j - 1) * hlayer.wt(1, 0) + olayer.val(i + 1, j - 1) * hlayer.wt(2, 0) + olayer.val(i + 1, j) * hlayer.wt(2, 1)
ElseIf i = Picture1.ScaleHeight - 1 And j = 0 Then
hlayer.val(i, j) = olayer.val(i, j + 1) * hlayer.wt(1, 2) + olayer.val(i - 1, j) * hlayer.wt(0, 1) + olayer.val(i - 1, j + 1) * hlayer.wt(0, 2)
ElseIf i = Picture1.ScaleHeight - 1 And j = Picture1.ScaleWidth - 1 Then
hlayer.val(i, j) = olayer.val(i, j - 1) * hlayer.wt(1, 0) + olayer.val(i - 1, j) * hlayer.wt(0, 1) + olayer.val(i - 1, j - 1) * hlayer.wt(0, 0)
ElseIf i = Picture1.ScaleHeight - 1 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
hlayer.val(i, j) = olayer.val(i, j + 1) * hlayer.wt(1, 2) + olayer.val(i, j - 1) * hlayer.wt(1, 0) + olayer.val(i - 1, j) * hlayer.wt(0, 1) + olayer.val(i - 1, j + 1) * hlayer.wt(0, 2) + olayer.val(i - 1, j - 1) * hlayer.wt(0, 0)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = 0 Then
hlayer.val(i, j) = olayer.val(i, j + 1) * hlayer.wt(1, 2) + olayer.val(i - 1, j) * hlayer.wt(0, 1) + olayer.val(i - 1, j + 1) * hlayer.wt(0, 2) + olayer.val(i + 1, j) * hlayer.wt(2, 1) + olayer.val(i + 1, j + 1) * hlayer.wt(2, 2)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = Picture1.ScaleWidth - 1 Then
hlayer.val(i, j) = olayer.val(i, j - 1) * hlayer.wt(1, 0) + olayer.val(i - 1, j) * hlayer.wt(0, 1) + olayer.val(i - 1, j - 1) * hlayer.wt(0, 0) + olayer.val(i + 1, j) * hlayer.wt(2, 1) + olayer.val(i + 1, j - 1) * hlayer.wt(2, 0)
Else
hlayer.val(i, j) = olayer.val(i, j - 1) * hlayer.wt(1, 0) + olayer.val(i, j + 1) * hlayer.wt(1, 2) + olayer.val(i + 1, j) * hlayer.wt(2, 1) + olayer.val(i + 1, j + 1) * hlayer.wt(2, 2) + olayer.val(i + 1, j - 1) * hlayer.wt(2, 0) + olayer.val(i - 1, j) * hlayer.wt(0, 1) + olayer.val(i - 1, j + 1) * hlayer.wt(0, 2) + olayer.val(i - 1, j - 1) * hlayer.wt(0, 0)
End If
hlayer.val(i, j) = 1 / (1 + Exp(-(hlayer.val(i, j) - 4)))
Picture2.PSet (X, Y), RGB(hlayer.val(i, j) * 255, hlayer.val(i, j) * 255, hlayer.val(i, j) * 255)
X = X + 1
Next
X = 0: Y = Y + 1
Next
End Sub
Private Sub hlayerTOolayer()
X = 0: Y = 0
For i = 0 To Picture1.ScaleHeight - 1
For j = 0 To Picture1.ScaleWidth - 1
If i = 0 And j = 0 Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2)
ElseIf i = 0 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0)
ElseIf i = 0 And j = Picture1.ScaleWidth - 1 Then
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0) + hlayer.val(i + 1, j) * olayer.wt(2, 1)
ElseIf i = Picture1.ScaleHeight - 1 And j = 0 Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2)
ElseIf i = Picture1.ScaleHeight - 1 And j = Picture1.ScaleWidth - 1 Then
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0)
ElseIf i = Picture1.ScaleHeight - 1 And (j > 0 Or j < Picture1.ScaleWidth - 1) Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = 0 Then
olayer.val(i, j) = hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2)
ElseIf (i > 0 Or i < Picture1.ScaleHeight - 1) And j = Picture1.ScaleWidth - 1 Then
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0)
Else
olayer.val(i, j) = hlayer.val(i, j - 1) * olayer.wt(1, 0) + hlayer.val(i, j + 1) * olayer.wt(1, 2) + hlayer.val(i + 1, j) * olayer.wt(2, 1) + hlayer.val(i + 1, j + 1) * olayer.wt(2, 2) + hlayer.val(i + 1, j - 1) * olayer.wt(2, 0) + hlayer.val(i - 1, j) * olayer.wt(0, 1) + hlayer.val(i - 1, j + 1) * olayer.wt(0, 2) + hlayer.val(i - 1, j - 1) * olayer.wt(0, 0)
End If
olayer.val(i, j) = 1 / (1 + Exp(-(olayer.val(i, j) - 4)))

If (maxerrtwo = 0) Then
maxerrtwo = (1 + Exp(-(olayer.val(i, j) - 4)))
Else
If (maxerrtwo < (1 + Exp(-(olayer.val(i, j) - 4)))) Then
maxerrtwo = (1 + Exp(-(olayer.val(i, j) - 4)))
End If
End If

Picture3.PSet (X, Y), RGB(olayer.val(i, j) * 255, olayer.val(i, j) * 255, olayer.val(i, j) * 255)
X = X + 1
Next
X = 0: Y = Y + 1
Next
End Sub

