VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "PACMAN BY DEViL"
   ClientHeight    =   12600
   ClientLeft      =   30
   ClientTop       =   1560
   ClientWidth     =   15930
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":424A
   ScaleHeight     =   840
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1062
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "PAUSE"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H8000000A&
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   9960
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "EXIT"
      Height          =   375
      Left            =   10800
      MaskColor       =   &H8000000A&
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   15480
      Picture         =   "Main.frx":56AEB
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   9000
      Left            =   4920
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   9000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   735
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE:"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tile(20, 20) '0=open 1=wall 2=pellets 3=super pellets
Dim tilea(20, 20)
Dim tileb(20, 20)
Dim a
Dim b
Dim d '1=right 2=down 3=left 4=up
Dim d2
Dim ani '0-4
Dim ga(4)
Dim gb(4)
Dim gd(4)
Dim gd2(4)
Dim gani(4)
Dim gc(4) '0=dead 1=alive

Dim score
Dim selecteda
Dim selectedb
Dim dying '0=alive #=death animation
Dim starting 'coundown ticks till ghost move
Dim super 'coundown till not super anymore



Private Sub Command1_Click()
End
'For z = 1 To 20
'For w = 1 To 20
'If tilea(z, w) = 1 And tileb(z, w) = 0 Then tile(z, w) = 2
'Next w
'Next z

'Open "map.txt" For Output As #1
'For z = 1 To 20
'For w = 1 To 20
'Print #1, tile(z, w)
'Print #1, tilea(z, w)
'Print #1, tileb(z, w)
'Next w
'Next z
'Close #1
'End
End Sub

Private Sub Command2_Click()
 If Command2.Caption = "RESUME" Then
        Command2.Caption = "PAUSE"
        Timer1.Enabled = True
    Else
        Command2.Caption = "RESUME"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Form_Load()

a = 10
b = 7
d = 4
d2 = 4
ani = 0

starting = 200

Picture1.Width = 30 * 19
Picture1.Height = 30 * 19
For z = 1 To 20
For w = 1 To 20
tile(z, w) = 2
tilea(z, w) = 1
tileb(z, w) = 0
Next w
Next z

Open "map.txt" For Input As #1
For z = 1 To 20
For w = 1 To 20
Input #1, tile(z, w)
Input #1, tilea(z, w)
Input #1, tileb(z, w)
Next w
Next z
Close #1

tile(1, 9) = 0
tile(2, 9) = 0
tile(18, 9) = 0
tile(19, 9) = 0
tile(2, 2) = 3: tilea(2, 2) = 2: tileb(2, 2) = 0
tile(18, 2) = 3: tilea(18, 2) = 2: tileb(18, 2) = 0
tile(2, 18) = 3: tilea(2, 18) = 2: tileb(2, 18) = 0
tile(18, 18) = 3: tilea(18, 18) = 2: tileb(18, 18) = 0

Call drawscreen

End Sub

Public Sub drawscreen()
For z = 1 To 20
For w = 1 To 20
Call drawtile(z, w)
Next w
Next z
End Sub

Public Sub drawtile(z, w)
Call Picture1.PaintPicture(Picture2.Image, (z - 1) * 30, (w - 1) * 30, 30, 30, tilea(z, w) * 30, tileb(z, w) * 30, 30, 30)
If z = selecteda And w = selectedb Then
    Picture1.Line ((z - 1) * 30, (w - 1) * 30)-((z) * 30 - 1, (w) * 30 - 1), , B
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then d2 = 3
If KeyCode = 38 Then d2 = 4
If KeyCode = 39 Then d2 = 1
If KeyCode = 40 Then d2 = 2
'1=right 2=down 3=left 4=up
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'selecteda = Int(X / 30) + 1
'selectedb = Int(Y / 30) + 1
'Call drawscreen
End Sub

Private Sub Picture1_Paint()
Call drawscreen
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'tempa = Int(X / 30) + 1
'tempb = Int(Y / 30) + 1
'Call drawscreen
tilea(selecteda, selectedb) = Int(X / 30)
tileb(selecteda, selectedb) = Int(Y / 30)
tile(selecteda, selectedb) = 1 'wall

Call drawtile(selecteda, selectedb)
End Sub

Private Sub Timer1_Timer()
If starting > 0 Then starting = starting - 1
'GHOST
If starting = 180 Then
    ga(1) = 10
    gb(1) = 7
    gd(1) = 4
    gani(1) = 0
    gc(1) = 1
End If
If starting = 160 Then
    ga(2) = 10
    gb(2) = 7
    gd(2) = 4
    gani(2) = 0
    gc(2) = 1
End If
If starting = 140 Then
    ga(3) = 10
    gb(3) = 7
    gd(3) = 4
    gani(3) = 0
    gc(3) = 1
End If
If starting = 120 Then
    ga(4) = 10
    gb(4) = 7
    gd(4) = 4
    gani(4) = 0
    gc(4) = 1
End If
'Stoping Game
If dying > 0 Then dying = dying + 0.5: If dying = 7 Then Timer1.Enabled = False: MsgBox "You're Dead!!": End

If super > 0 Then super = super - 1

If dying = 0 Then
    ani = ani + 1
    If ani = 4 Then ani = 0
End If
'Super Pellet Animation Blink

If tile(2, 2) = 3 And (ani = 0 Or ani = 1) Then tilea(2, 2) = 2
If tile(2, 2) = 3 And (ani = 2 Or ani = 3) Then tilea(2, 2) = 3

If tile(18, 2) = 3 And (ani = 0 Or ani = 1) Then tilea(18, 2) = 2
If tile(18, 2) = 3 And (ani = 2 Or ani = 3) Then tilea(18, 2) = 3

If tile(2, 18) = 3 And (ani = 0 Or ani = 1) Then tilea(2, 18) = 2
If tile(2, 18) = 3 And (ani = 2 Or ani = 3) Then tilea(2, 18) = 3

If tile(18, 18) = 3 And (ani = 0 Or ani = 1) Then tilea(18, 18) = 2
If tile(18, 18) = 3 And (ani = 2 Or ani = 3) Then tilea(18, 18) = 3

'Super Pellet



'1=right 2=down 3=left 4=up
If dying = 0 Then
    If ani = 0 And d = 1 Then a = a + 1
    If ani = 0 And d = 2 Then b = b + 1
    If ani = 0 And d = 3 Then a = a - 1
    If ani = 0 And d = 4 Then b = b - 1
End If

'Ghost
'Ghost Directions AI
If super = 0 Then
If a > ga(1) Then gd2(1) = 1
If a < ga(1) Then gd2(1) = 3
If b > gb(1) Then gd2(1) = 2
If b < gb(1) Then gd2(1) = 4

If a < ga(2) Then gd2(2) = 3
If a > ga(2) Then gd2(2) = 1
If b < gb(2) Then gd2(2) = 4
If b > gb(2) Then gd2(2) = 2

If b > gb(3) Then gd2(3) = 2
If b < gb(3) Then gd2(3) = 4
If a > ga(3) Then gd2(3) = 1
If a < ga(3) Then gd2(3) = 3

If b < gb(4) Then gd2(4) = 4
If b > gb(4) Then gd2(4) = 2
If a < ga(4) Then gd2(4) = 3
If a > ga(4) Then gd2(4) = 1
End If
'Ghost Directions AI
'Super Pallet Effect
If super = 1 Then
If a > ga(1) Then gd2(1) = 3
If a < ga(1) Then gd2(1) = 1
If b > gb(1) Then gd2(1) = 4
If b < gb(1) Then gd2(1) = 2

If a < ga(2) Then gd2(2) = 1
If a > ga(2) Then gd2(2) = 3
If b < gb(2) Then gd2(2) = 2
If b > gb(2) Then gd2(2) = 4

If b > gb(3) Then gd2(3) = 4
If b < gb(3) Then gd2(3) = 2
If a > ga(3) Then gd2(3) = 3
If a < ga(3) Then gd2(3) = 1

If b < gb(4) Then gd2(4) = 2
If b > gb(4) Then gd2(4) = 4
If a < ga(4) Then gd2(4) = 1
If a > ga(4) Then gd2(4) = 3
End If
'Super Pallet Effect


'tile(1, 9) = 0
'tile(2, 9) = 0
'tile(18, 9) = 0
'tile(19, 9) = 0
If dying = 0 Then
For z = 1 To 4
    If ani = 0 And gd(z) = 1 Then ga(z) = ga(z) + 1
    If ani = 0 And gd(z) = 2 Then gb(z) = gb(z) + 1
    If ani = 0 And gd(z) = 3 Then ga(z) = ga(z) - 1
    If ani = 0 And gd(z) = 4 Then gb(z) = gb(z) - 1
Next z
End If

'Portal
If a = 0 And b = 9 Then a = 19: d2 = 3
If a = 20 And b = 9 Then a = 1: d2 = 1
For z = 1 To 4
If ga(z) = 0 And gb(z) = 9 Then ga(z) = 19: gd2(z) = 3
If ga(z) = 20 And gb(z) = 9 Then ga(z) = 1: gd2(z) = 1
Next z
'Portal
'Dying
If dying = 0 And super = 0 Then
    For z = 1 To 4
    If ga(z) = a And gb(z) = b Then dying = 1
    If ani = 2 Then
    If ga(z) = a + 1 And d = 1 And gd(z) = 3 And gb(z) = b Then dying = 1
    If ga(z) = a - 1 And d = 3 And gd(z) = 1 And gb(z) = b Then dying = 1
    If gb(z) = b + 1 And d = 2 And gd(z) = 4 And ga(z) = a Then dying = 1
    If gb(z) = b - 1 And d = 4 And gd(z) = 2 And ga(z) = a Then dying = 1
    End If
    Next z
End If
'SCORE

If ani = 0 And tile(a, b) = 3 Then tile(a, b) = 0: tilea(a, b) = 0: super = 100

If ani = 0 And tile(a, b) = 2 Then
    score = score + 1: tile(a, b) = 0: tilea(a, b) = 0
    ok = 0
    For z = 1 To 19
    For w = 1 To 19
    If tile(z, w) = 2 Or tile(z, w) = 3 Then ok = 1
    Next w
    Next z
    If ok = 0 Then Timer1.Enabled = False: MsgBox "!!!YOU WIN!!!": End
End If
'SCORE
Label1.Caption = score
'Nested
If ani = 0 Then
    If d2 = 1 And tile(a + 1, b) <> 1 Then d = d2
    If d2 = 2 And tile(a, b + 1) <> 1 Then d = d2
    If d2 = 3 And tile(a - 1, b) <> 1 Then d = d2
    If d2 = 4 And tile(a, b - 1) <> 1 Then d = d2
End If

If ani = 0 Then 'GHOSTS
For z = 1 To 4
    If ga(z) <> 0 Then
    If gd2(z) = 1 And tile(ga(z) + 1, gb(z)) <> 1 Then gd(z) = gd2(z)
    If gd2(z) = 2 And tile(ga(z), gb(z) + 1) <> 1 Then gd(z) = gd2(z)
    If gd2(z) = 3 And tile(ga(z) - 1, gb(z)) <> 1 Then gd(z) = gd2(z)
    If gd2(z) = 4 And tile(ga(z), gb(z) - 1) <> 1 Then gd(z) = gd2(z)
    End If
Next z
End If

If ani = 0 Then
    If d = 1 And tile(a + 1, b) = 1 Then
        If tile(a, b - 1) <> 1 Then d = 4: d2 = d Else d = 2: d2 = d
    End If
    If d = 3 And tile(a - 1, b) = 1 Then
        If tile(a, b - 1) <> 1 Then d = 4: d2 = d Else d = 2: d2 = d
    End If
    If d = 2 And tile(a, b + 1) = 1 Then
        If tile(a - 1, b) <> 1 Then d = 3: d2 = d Else d = 1: d2 = d
    End If
    If d = 4 And tile(a, b - 1) = 1 Then
        If tile(a - 1, b) <> 1 Then d = 3: d2 = d Else d = 1: d2 = d
    End If
End If

If ani = 0 Then 'GHOSTS
For z = 1 To 4
    If ga(z) <> 0 Then
    If gd(z) = 1 And tile(ga(z) + 1, gb(z)) = 1 Then
        If tile(ga(z), gb(z) - 1) <> 1 Then gd(z) = 4: gd2(z) = gd(z) Else gd(z) = 2: gd2(z) = gd(z)
    End If
    If gd(z) = 3 And tile(ga(z) - 1, gb(z)) = 1 Then
        If tile(ga(z), gb(z) - 1) <> 1 Then gd(z) = 4: gd2(z) = gd(z) Else gd(z) = 2: gd2(z) = gd(z)
    End If
    If gd(z) = 2 And tile(ga(z), gb(z) + 1) = 1 Then
        If tile(ga(z) - 1, gb(z)) <> 1 Then gd(z) = 3: gd2(z) = gd(z) Else gd(z) = 1: gd2(z) = gd(z)
    End If
    If gd(z) = 4 And tile(ga(z), gb(z) - 1) = 1 Then
        If tile(ga(z) - 1, gb(z)) <> 1 Then gd(z) = 3: gd2(z) = gd(z) Else gd(z) = 1: gd2(z) = gd(z)
    End If
    End If
Next z
End If
'BIG BUG
Call drawtile(a, b)
'BIG BUG

'If d = 4 Then Call drawtile(a, b - 1)
'If d = 2 Then Call drawtile(a, b + 1)
'If d = 1 Then Call drawtile(a - 1, b)
'If d = 3 Then Call drawtile(a + 1, b)
Call drawscreen
'1=right 2=down 3=left 4=up
'Pacman Drawing
If dying = 0 Then
If d = 1 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30) + (ani * 7.5), ((b - 1) * 30), 30, 30, (d - 1) * 30, (2 + ani + 4) * 30, 30, 30, vbMergePaint)
If d = 2 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30), ((b - 1) * 30) + (ani * 7.5), 30, 30, (d - 1) * 30, (2 + ani + 4) * 30, 30, 30, vbMergePaint)
If d = 3 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30) - (ani * 7.5), ((b - 1) * 30), 30, 30, (d - 1) * 30, (2 + ani + 4) * 30, 30, 30, vbMergePaint)
If d = 4 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30), ((b - 1) * 30) - (ani * 7.5), 30, 30, (d - 1) * 30, (2 + ani + 4) * 30, 30, 30, vbMergePaint)
If d = 1 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30) + (ani * 7.5), ((b - 1) * 30), 30, 30, (d - 1) * 30, (2 + ani) * 30, 30, 30, vbSrcAnd)
If d = 2 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30), ((b - 1) * 30) + (ani * 7.5), 30, 30, (d - 1) * 30, (2 + ani) * 30, 30, 30, vbSrcAnd)
If d = 3 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30) - (ani * 7.5), ((b - 1) * 30), 30, 30, (d - 1) * 30, (2 + ani) * 30, 30, 30, vbSrcAnd)
If d = 4 Then Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30), ((b - 1) * 30) - (ani * 7.5), 30, 30, (d - 1) * 30, (2 + ani) * 30, 30, 30, vbSrcAnd)
End If

If dying > 0 Then
    Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30), ((b - 1) * 30), 30, 30, Int(dying - 1) * 30, (11) * 30, 30, 30, vbMergePaint)
    Call Picture1.PaintPicture(Picture2.Image, ((a - 1) * 30), ((b - 1) * 30), 30, 30, Int(dying - 1) * 30, (10) * 30, 30, 30, vbSrcAnd)
End If

'Ghost Draw
If dying = 0 Then
If super = 0 Then
    For z = 1 To 4
    If gd(z) = 1 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) + (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1 + 4) * 30, (2 + 4 + (z - 1)) * 30, 30, 30, vbMergePaint)
    If gd(z) = 2 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) + (ani * 7.5), 30, 30, (gd(z) - 1 + 4) * 30, (2 + 4 + (z - 1)) * 30, 30, 30, vbMergePaint)
    If gd(z) = 3 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) - (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1 + 4) * 30, (2 + 4 + (z - 1)) * 30, 30, 30, vbMergePaint)
    If gd(z) = 4 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) - (ani * 7.5), 30, 30, (gd(z) - 1 + 4) * 30, (2 + 4 + (z - 1)) * 30, 30, 30, vbMergePaint)
    If gd(z) = 1 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) + (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1 + 4) * 30, (2 + (z - 1)) * 30, 30, 30, vbSrcAnd)
    If gd(z) = 2 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) + (ani * 7.5), 30, 30, (gd(z) - 1 + 4) * 30, (2 + (z - 1)) * 30, 30, 30, vbSrcAnd)
    If gd(z) = 3 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) - (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1 + 4) * 30, (2 + (z - 1)) * 30, 30, 30, vbSrcAnd)
    If gd(z) = 4 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) - (ani * 7.5), 30, 30, (gd(z) - 1 + 4) * 30, (2 + (z - 1)) * 30, 30, 30, vbSrcAnd)
    Next z
End If
If super > 0 Then
    For z = 1 To 4
    If gd(z) = 1 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) + (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1) * 30, (13) * 30, 30, 30, vbMergePaint)
    If gd(z) = 2 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) + (ani * 7.5), 30, 30, (gd(z) - 1) * 30, (13) * 30, 30, 30, vbMergePaint)
    If gd(z) = 3 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) - (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1) * 30, (13) * 30, 30, 30, vbMergePaint)
    If gd(z) = 4 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) - (ani * 7.5), 30, 30, (gd(z) - 1) * 30, (13) * 30, 30, 30, vbMergePaint)
    If gd(z) = 1 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) + (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1) * 30, (12) * 30, 30, 30, vbSrcAnd)
    If gd(z) = 2 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) + (ani * 7.5), 30, 30, (gd(z) - 1) * 30, (12) * 30, 30, 30, vbSrcAnd)
    If gd(z) = 3 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30) - (ani * 7.5), ((gb(z) - 1) * 30), 30, 30, (gd(z) - 1) * 30, (12) * 30, 30, 30, vbSrcAnd)
    If gd(z) = 4 Then Call Picture1.PaintPicture(Picture2.Image, ((ga(z) - 1) * 30), ((gb(z) - 1) * 30) - (ani * 7.5), 30, 30, (gd(z) - 1) * 30, (12) * 30, 30, 30, vbSrcAnd)
    Next z
End If
End If

End Sub
